// Copyright 2021 Tamas Gulacsi. All rights reserved.

package main

import (
	"context"
	"encoding/hex"
	"flag"
	"fmt"
	"log/slog"
	"math"
	"os"
	"os/signal"
	"strings"
	"syscall"
	"time"

	"github.com/UNO-SOFT/spreadsheet"
	"github.com/UNO-SOFT/zlog/v2"
	"github.com/peterbourgon/ff/v3/ffcli"

	"github.com/johnfercher/maroto/v2"
	"github.com/johnfercher/maroto/v2/pkg/components/text"
	"github.com/johnfercher/maroto/v2/pkg/config"
	"github.com/johnfercher/maroto/v2/pkg/consts/align"
	"github.com/johnfercher/maroto/v2/pkg/consts/orientation"
	"github.com/johnfercher/maroto/v2/pkg/consts/pagesize"
	"github.com/johnfercher/maroto/v2/pkg/core"
	"github.com/johnfercher/maroto/v2/pkg/props"
)

var verbose zlog.VerboseVar
var logger = zlog.NewLogger(zlog.MaybeConsoleHandler(&verbose, os.Stderr)).SLog()

func main() {
	if err := Main(); err != nil {
		slog.Error("MAIN", "error", err)
		os.Exit(1)
	}
}

func Main() error {
	fgColor := props.Color{}
	bgColor := props.Color{Red: 255, Green: 255, Blue: 255}
	alternateFgColor := fgColor
	alternateBgColor := Color{Color: props.Color{
		Red:   230,
		Green: 230,
		Blue:  230,
	}}

	fs := flag.NewFlagSet("csv2pdf", flag.ContinueOnError)
	fs.Var(&verbose, "v", "logging verbosity")
	flagEnc := fs.String("charset", spreadsheet.EncName, "csv charset name")
	flagOut := fs.String("o", "", "output file name (default input file + .pdf)")
	flagColor := fs.String("alternate-color", alternateBgColor.String(), "alternate background color")
	flagLandscape := fs.Bool("L", false, "landscape orientation (default: portrait)")
	flagFontSize := fs.Float64("f", 8, "font size")
	flagPrintPagenum := fs.Bool("print-pagenum", false, "print page numbers")

	app := ffcli.Command{Name: "csv2pdf", FlagSet: fs,
		Exec: func(ctx context.Context, args []string) error {
			var inp string
			if len(args) != 0 {
				inp = args[0]
			}
			cr, err := spreadsheet.OpenCsv(inp, *flagEnc)
			if err != nil {
				return err
			}
			defer cr.Close()

			headers, err := cr.Read()
			if err != nil {
				return err
			}
			widths := make([]float64, len(headers))
			var avg float64
			for i, s := range headers {
				widths[i] = float64(len(s))
				avg += widths[i]
			}
			slog.Debug("widths", "headers", headers, "widths", widths)

			contents, err := cr.ReadAll()
			if err != nil {
				return err
			}
			for _, row := range contents {
				for i, s := range row {
					widths[i] += float64(len(s))
					avg += widths[i]
				}
			}
			avg /= float64(len(widths))
			avg /= float64(len(contents) + 1)
			gridSize := make([]uint, len(headers))
			for i, w := range widths {
				gridSize[i] = uint(math.Round(w / avg / 4))
				if gridSize[i] == 0 {
					gridSize[i] = 1
				}
			}
			orient := orientation.Vertical
			if *flagLandscape {
				orient = orientation.Horizontal
			}
			cfg := config.NewBuilder().
				WithCompression(true).
				WithCreationDate(time.Now()).
				WithCreator("UNO-SOFT/spreadsheet/csv2pdf", false).
				WithMargins(1, 1, 1).
				WithPageSize(pagesize.A4).
				WithOrientation(orient).
				WithSubject(inp, true)
			if *flagPrintPagenum {
				cfg = cfg.WithPageNumber("{current}/{total}", props.RightBottom)
			}
			m := maroto.NewMetricsDecorator(maroto.New(cfg.Build()))

			headerProps := props.Text{Family: "arial", Size: *flagFontSize * 1.375, Align: align.Center, Color: &alternateFgColor}
			headerCols := make([]core.Col, 0, len(headers))
			for _, h := range headers {
				headerCols = append(headerCols, text.NewCol(1, h, headerProps))
			}
			m.AddRow(headerProps.Size*7/8, headerCols...).
				WithStyle(&props.Cell{BackgroundColor: &alternateBgColor.Color})

			type rowProp struct {
				props.Text
				props.Cell
			}
			contentProps := []rowProp{
				rowProp{
					Text: props.Text{Family: "courier", Size: *flagFontSize, Color: &fgColor},
					Cell: props.Cell{BackgroundColor: &bgColor},
				},
				rowProp{
					Text: props.Text{Family: "courier", Size: *flagFontSize, Color: &alternateFgColor},
					Cell: props.Cell{BackgroundColor: &alternateBgColor.Color},
				},
			}
			if len(contents) != 0 {
				content := make([]core.Col, len(contents[0]))
				for i, cc := range contents {
					props := &contentProps[i%2]
					for i, c := range cc {
						content[i] = text.NewCol(1,
							c, props.Text)
					}
					m.AddRow(contentProps[0].Text.Size, content...).
						WithStyle(&props.Cell)
				}
			}

			doc, err := m.Generate()
			if err != nil {
				return err
			}

			out := *flagOut
			if out == "" &&
				len(args) != 0 && args[0] != "" && args[0] != "-" {
				out = args[0] + ".pdf"
			}
			if out == "" || out == "-" {
				_, err = os.Stdout.Write(doc.GetBytes())
				return err
			}
			return doc.Save(out)
		},
	}

	args := make([]string, 0, len(os.Args))
	for _, a := range os.Args[1:] {
		if strings.HasPrefix(a, "-f") && len(a) > 2 && '0' <= a[2] && a[2] <= '9' {
			args = append(args, "-f", a[2:])
		} else {
			args = append(args, a)
		}
	}
	slog.Debug("args", "original", os.Args[1:], "fixed", args)
	if err := app.Parse(args); err != nil {
		return err
	}

	if *flagColor != "" {
		if err := alternateBgColor.Parse(*flagColor); err != nil {
			return err
		}
		// 0.2989 R + 0.5870 G + 0.1140 B
		if 0.2989*float64(alternateBgColor.Red)+
			0.5870*float64(alternateBgColor.Green)+
			0.1140*float64(alternateBgColor.Blue) > 127 {
			alternateFgColor = props.Color{}
		} else {
			alternateFgColor = props.Color{Red: 255, Green: 255, Blue: 255}
		}
	}

	ctx, cancel := signal.NotifyContext(context.Background(),
		os.Interrupt, syscall.SIGTERM)
	defer cancel()
	return app.Run(ctx)
}

type Color struct {
	props.Color
}

func (c *Color) String() string {
	return fmt.Sprintf("%x%x%x", c.Red, c.Green, c.Blue)
}
func (c *Color) Parse(s string) error {
	b, err := hex.DecodeString(s)
	if err != nil {
		return err
	}
	c.Red, c.Green, c.Blue = int(b[0]), int(b[1]), int(b[2])
	return nil
}
