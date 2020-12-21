package main

import (
	"bufio"
	"encoding/csv"
	"flag"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"unicode"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/htmlindex"

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
	encName := os.Getenv("LANG")
	if i := strings.IndexByte(encName, '.'); i >= 0 {
		encName = strings.ToUpper(encName[i+1:])
	}
	if encName == "" {
		encName = "UTF-8"
	}
	flag.StringVar(&encName, "charset", encName, "csv charset name")
	flag.Parse()

	encName = strings.ToUpper(encName)
	var enc encoding.Encoding
	if !(encName == "" || encName == "UTF-8" || encName == "UTF8") {
		var err error
		if enc, err = htmlindex.Get(encName); err != nil {
			return fmt.Errorf("%q: %w", encName, err)
		}
	}
	log.Printf("encoding: %s", enc)

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
		if err := copyFile(w, sheetName, enc, fn); err != nil {
			return fmt.Errorf("%q: %w", fn, err)
		}
	}

	if err := w.Close(); err != nil {
		return err
	}
	return fh.Close()
}

func copyFile(w spreadsheet.Writer, sheetName string, enc encoding.Encoding, fn string) error {
	fh := os.Stdin
	var err error
	if !(fn == "" || fn == "-") {
		if fh, err = os.Open(fn); err != nil {
			return fmt.Errorf("open %q: %w", fn, err)
		}
	}
	defer fh.Close()
	r := io.Reader(fh)
	if enc != nil {
		r = enc.NewDecoder().Reader(r)
	}
	br := bufio.NewReaderSize(r, 1<<20)
	b, err := br.Peek(1024)
	if err != nil && len(b) == 0 {
		return err
	}
	var sep = rune(',')
	for _, r := range string(b) {
		if r == '"' || r == '_' || unicode.IsLetter(r) || unicode.IsNumber(r) {
			continue
		}
		sep = r
		break
	}

	cr := csv.NewReader(br)
	cr.ReuseRecord = true
	cr.Comma = sep

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
			rowI = append(rowI, s)
		}
		sheet.AppendRow(rowI...)
	}
	return sheet.Close()
}
