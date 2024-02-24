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

	"github.com/UNO-SOFT/spreadsheet"
	"github.com/UNO-SOFT/zlog/v2"
	"github.com/peterbourgon/ff/v3/ffcli"

	"github.com/johnfercher/maroto/pkg/color"
	"github.com/johnfercher/maroto/pkg/consts"
	"github.com/johnfercher/maroto/pkg/pdf"
	"github.com/johnfercher/maroto/pkg/props"
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
	alternateColor := Color{Color: color.Color{
		Red:   230,
		Green: 230,
		Blue:  230,
	}}

	fs := flag.NewFlagSet("csv2pdf", flag.ContinueOnError)
	fs.Var(&verbose, "v", "logging verbosity")
	flagEnc := fs.String("charset", spreadsheet.EncName, "csv charset name")
	flagOut := fs.String("o", "", "output file name (default input file + .pdf)")
	flagColor := fs.String("alternate-color", alternateColor.String(), "alternate color")
	flagLandscape := fs.Bool("L", false, "landscape orientation (default: portrait)")
	flagFontSize := fs.Float64("f", 8, "font size")
	flagPrintPagenum := fs.Bool("print-pagenum", false, "print page numbers")

	app := ffcli.Command{Name: "csv2pdf", FlagSet: fs,
		Exec: func(ctx context.Context, args []string) error {
			cr, err := spreadsheet.OpenCsv(args[0], *flagEnc)
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
			orientation := consts.Portrait
			if *flagLandscape {
				orientation = consts.Landscape
			}
			m := pdf.NewMaroto(orientation, consts.A4)
			if *flagPrintPagenum {
				//
			}
			m.TableList(headers, contents, props.TableList{
				HeaderProp: props.TableListContent{
					Family:    consts.Arial,
					Style:     consts.Bold,
					Size:      *flagFontSize * 1.375,
					GridSizes: gridSize,
				},
				ContentProp: props.TableListContent{
					Family:    consts.Courier,
					Style:     consts.Normal,
					Size:      *flagFontSize,
					GridSizes: gridSize,
				},
				Align:                consts.Center,
				AlternatedBackground: &alternateColor.Color,
				HeaderContentSpace:   *flagFontSize * 1.2,
				Line:                 false,
			})
			out := *flagOut
			if out == "" &&
				len(args) != 0 && args[0] != "" && args[0] != "-" {
				out = args[0] + ".pdf"
			}
			if out == "" || out == "-" {
				buf, err := m.Output()
				if err != nil {
					return err
				}
				_, err = os.Stdout.Write(buf.Bytes())
				return err
			}
			return m.OutputFileAndClose(out)
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

	if err := alternateColor.Parse(*flagColor); err != nil {
		return err
	}

	ctx, cancel := signal.NotifyContext(context.Background(),
		os.Interrupt, syscall.SIGTERM)
	defer cancel()
	return app.Run(ctx)
}

type Color struct {
	color.Color
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
