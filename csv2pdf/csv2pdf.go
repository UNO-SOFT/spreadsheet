package main

import (
	"encoding/hex"
	"flag"
	"fmt"
	"log"
	"math"
	"os"

	"github.com/UNO-SOFT/spreadsheet"

	"github.com/johnfercher/maroto/pkg/color"
	"github.com/johnfercher/maroto/pkg/consts"
	"github.com/johnfercher/maroto/pkg/pdf"
	"github.com/johnfercher/maroto/pkg/props"
)

func main() {
	if err := Main(); err != nil {
		log.Fatalf("%+v", err)
	}
}

func Main() error {
	alternateColor := Color{Color: color.Color{
		Red:   230,
		Green: 230,
		Blue:  230,
	}}
	flagEnc := flag.String("charset", spreadsheet.EncName, "csv charset name")
	flagOut := flag.String("o", "", "output file name (default stdout")
	flagColor := flag.String("alternate-color", alternateColor.String(), "alternate color")
	flag.Parse()

	if err := alternateColor.Parse(*flagColor); err != nil {
		return err
	}

	cr, err := spreadsheet.OpenCsv(flag.Arg(0), *flagEnc)
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
	m := pdf.NewMaroto(consts.Landscape, consts.A4)
	m.TableList(headers, contents, props.TableList{
		HeaderProp: props.TableListContent{
			Family:    consts.Arial,
			Style:     consts.Bold,
			Size:      11.0,
			GridSizes: gridSize,
		},
		ContentProp: props.TableListContent{
			Family:    consts.Courier,
			Style:     consts.Normal,
			Size:      8.0,
			GridSizes: gridSize,
		},
		Align:                consts.Center,
		AlternatedBackground: &alternateColor.Color,
		HeaderContentSpace:   10.0,
		Line:                 false,
	})
	if *flagOut == "" || *flagOut == "-" {
		buf, err := m.Output()
		if err != nil {
			return err
		}
		_, err = os.Stdout.Write(buf.Bytes())
		return err
	}
	return m.OutputFileAndClose(*flagOut)
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
	return err
}
