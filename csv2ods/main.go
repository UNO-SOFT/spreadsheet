package main

import (
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/UNO-SOFT/spreadsheet"
	"github.com/UNO-SOFT/spreadsheet/ods"
	"github.com/UNO-SOFT/spreadsheet/xlsx"
)

func main() {
	if err := Main(); err != nil {
		log.Fatal(err)
	}
}

func Main() error {
	flagEnc := flag.String("charset", spreadsheet.EncName, "csv charset name")
	flag.Parse()

	fn := flag.Arg(0)
	fh := os.Stdout
	if !(fn == "" || fn == "-") {
		var err error
		if fh, err = os.Create(fn); err != nil {
			return err
		}
	}
	defer fh.Close()
	var w spreadsheet.Writer
	if strings.HasSuffix(fn, ".xlsx") {
		w = xlsx.NewWriter(fh)
	} else {
		var err error
		if w, err = ods.NewWriter(fh); err != nil {
			return err
		}
	}

	for i, fn := range flag.Args()[1:] {
		sheetName := fmt.Sprintf("Sheet%d", i+1)
		if i := strings.IndexByte(fn, ':'); i >= 0 {
			sheetName, fn = fn[:i], fn[i+1:]
		} else if fn != "" && fn != "-" {
			sheetName = strings.TrimSuffix(filepath.Base(fn), ".csv")
		}
		if err := copyFile(w, sheetName, fn, *flagEnc); err != nil {
			return fmt.Errorf("%q: %w", fn, err)
		}
	}

	if err := w.Close(); err != nil {
		return err
	}
	return fh.Close()
}

func copyFile(w spreadsheet.Writer, sheetName string, fn, encName string) error {
	cr, err := spreadsheet.OpenCsv(fn, encName)
	if err != nil {
		return err
	}
	defer cr.Close()

	row, err := cr.Read()
	if err != nil {
		return err
	}
	cols := make([]spreadsheet.Column, len(row))
	for i, r := range row {
		cols[i].Name = r
		cols[i].Header.FontBold = true
	}
	sheet, err := w.NewSheet(sheetName, cols)
	if err != nil {
		return err
	}

	var rowI []interface{}
	for {
		if row, err = cr.Read(); err != nil {
			if err == io.EOF {
				break
			}
			return err
		}
		rowI = rowI[:0]
		for _, s := range row {
			if len(s) != 0 {
				if len(s) == 10 && likelyDate(s) {
					if t, err := time.ParseInLocation("2006-01-02", s, time.Local); err == nil {
						rowI = append(rowI, t)
						continue
					}
				}
				if likelyNumber(s) {
					if i, err := strconv.ParseInt(s, 10, 64); err == nil {
						rowI = append(rowI, i)
						continue
					}
					if f, err := strconv.ParseFloat(s, 64); err == nil {
						rowI = append(rowI, f)
						continue
					}
				}
			}
			rowI = append(rowI, s)
		}
		sheet.AppendRow(rowI...)
	}
	return sheet.Close()
}

func likelyNumber(s string) bool {
	return strings.IndexFunc(s,
		func(r rune) bool {
			return !(r == '-' || r == '.' || '0' <= r && r <= '9')
		}) < 0
}
func likelyDate(s string) bool {
	return strings.IndexFunc(s,
		func(r rune) bool {
			return !(r == '-' || '0' <= r && r <= '9')
		}) < 0
}
