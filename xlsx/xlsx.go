// Copyright 2020, Tamás Gulácsi.
//
// SPDX-License-Identifier: Apache-2.0

package xlsx

import (
	"fmt"
	"io"
	"strings"
	"sync"
	"time"

	errors "golang.org/x/xerrors"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/UNO-SOFT/spreadsheet"
)

var _ = (spreadsheet.Writer)((*XLSXWriter)(nil))

type XLSXWriter struct {
	mu     sync.Mutex
	xl     *excelize.File
	w      io.Writer
	sheets []string
	styles map[string]int
}

type XLSXSheet struct {
	Name string

	mu  sync.Mutex
	xl  *excelize.File
	row int64
}

// NewWriter returns a new spreadsheet.Writer.
//
// This writer allows concurrent writes to separate sheets.
//
// This writer collects everything in memory, so big sheets may impose problems.
func NewWriter(w io.Writer) *XLSXWriter {
	return &XLSXWriter{w: w, xl: excelize.NewFile()}
}

func (xlw *XLSXWriter) Close() error {
	if xlw == nil {
		return nil
	}
	xlw.mu.Lock()
	defer xlw.mu.Unlock()
	xl, w := xlw.xl, xlw.w
	xlw.xl, xlw.w = nil, nil
	if xl == nil || w == nil {
		return nil
	}
	_, err := xl.WriteTo(w)
	return err
}
func (xlw *XLSXWriter) NewSheet(name string, columns []spreadsheet.Column) (spreadsheet.Sheet, error) {
	xlw.mu.Lock()
	defer xlw.mu.Unlock()
	xlw.sheets = append(xlw.sheets, name)
	if len(xlw.sheets) == 1 { // first
		xlw.xl.SetSheetName("Sheet1", name)
	} else {
		xlw.xl.NewSheet(name)
	}
	var hasHeader bool
	for i, c := range columns {
		col, err := excelize.ColumnNumberToName(i + 1)
		if err != nil {
			return nil, err
		}
		if s := xlw.getStyle(c.Column); s != 0 {
			if err = xlw.xl.SetColStyle(name, col, s); err != nil {
				return nil, err
			}
		}
		if s := xlw.getStyle(c.Header); s != 0 {
			if err = xlw.xl.SetCellStyle(name, col+"1", col+"1", s); err != nil {
				return nil, err
			}
		}
		if c.Name != "" {
			hasHeader = true
			if err = xlw.xl.SetCellStr(name, col+"1", c.Name); err != nil {
				return nil, err
			}
		}
	}
	xls := &XLSXSheet{xl: xlw.xl, Name: name}
	if hasHeader {
		xls.row++
	}
	return xls, nil
}

func (xlw *XLSXWriter) getStyle(style spreadsheet.Style) int {
	if !style.FontBold && style.Format == "" {
		return 0
	}
	xlw.mu.Lock()
	defer xlw.mu.Unlock()
	k := fmt.Sprintf("%t\t%s", style.FontBold, style.Format)
	s, ok := xlw.styles[k]
	if ok {
		return s
	}
	var buf strings.Builder
	buf.WriteByte('{')
	if style.FontBold {
		buf.WriteString(`"font":{"bold":true}`)
	}
	if style.Format != "" {
		if buf.Len() > 1 {
			buf.WriteByte(',')
		}
		fmt.Fprintf(&buf, `"custom_number_format":%q`, style.Format)
	}
	buf.WriteByte('}')
	s, err := xlw.xl.NewStyle(buf.String())
	if err != nil {
		panic(err)
	}
	if xlw.styles == nil {
		xlw.styles = make(map[string]int)
	}
	xlw.styles[k] = s
	return s
}

func (xls *XLSXSheet) Close() error { return nil }
func (xls *XLSXSheet) AppendRow(values ...interface{}) error {
	xls.mu.Lock()
	defer xls.mu.Unlock()
	xls.row++
	for i, v := range values {
		axis, err := excelize.CoordinatesToCellName(i+1, int(xls.row))
		if err != nil {
			return errors.Errorf("%d/%d: %w", i, int(xls.row), err)
		}
		isNil := v == nil
		if !isNil {
			if t, ok := v.(time.Time); ok {
				if isNil = t.IsZero(); !isNil {
					if err = xls.xl.SetCellStr(xls.Name, axis, t.Format("2006-01-02")); err != nil {
						return errors.Errorf("%s[%s]: %w", xls.Name, axis, err)
					}
					continue
				}
			}
		}
		if isNil {
			continue
		}
		if err = xls.xl.SetCellValue(xls.Name, axis, v); err != nil {
			return errors.Errorf("%s[%s]: %w", xls.Name, axis, err)
		}
	}
	return nil
}
