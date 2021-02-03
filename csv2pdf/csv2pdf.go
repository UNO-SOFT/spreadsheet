package main

import (
	"flag"
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
	flagEnc := flag.String("charset", spreadsheet.EncName, "csv charset name")
	flagOut := flag.String("o", "", "output file name (default stdout")
	flag.Parse()

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
		Align: consts.Center,
		AlternatedBackground: &color.Color{
			Red:   100,
			Green: 120,
			Blue:  255,
		},
		HeaderContentSpace: 10.0,
		Line:               false,
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
