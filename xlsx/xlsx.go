// Copyright 2020, 2023 Tamás Gulácsi.
//
// SPDX-License-Identifier: Apache-2.0

package xlsx

import (
	"database/sql"
	"database/sql/driver"
	"fmt"
	"io"
	"sync"
	"time"

	"github.com/UNO-SOFT/spreadsheet"
	"github.com/xuri/excelize/v2"
)

var _ = (spreadsheet.Writer)((*XLSXWriter)(nil))

type XLSXWriter struct {
	w      io.Writer
	xl     *excelize.File
	styles map[string]int
	sheets []string
	mu     sync.Mutex
}

type XLSXSheet struct {
	xl   *excelize.File
	Name string
	row  int64
	mu   sync.Mutex
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
	k := fmt.Sprintf("%t\t%s", style.FontBold, style.Format)
	s, ok := xlw.styles[k]
	if ok {
		return s
	}
	var st excelize.Style
	if style.FontBold {
		st.Font = &excelize.Font{Bold: true}
	}
	if style.Format != "" {
		st.CustomNumFmt = &style.Format
	}
	s, err := xlw.xl.NewStyle(&st)
	if err != nil {
		panic(err)
	}
	if xlw.styles == nil {
		xlw.styles = make(map[string]int)
	}
	xlw.styles[k] = s
	return s
}

// MaxRowCount is the number of maximum rows.
const MaxRowCount = 1_048_576

func (xls *XLSXSheet) Close() error { return nil }
func (xls *XLSXSheet) AppendRow(values ...interface{}) error {
	xls.mu.Lock()
	defer xls.mu.Unlock()
	if xls.row >= MaxRowCount {
		return spreadsheet.ErrTooManyRows
	}
	xls.row++
	for i, v := range values {
		axis, err := excelize.CoordinatesToCellName(i+1, int(xls.row))
		if err != nil {
			return fmt.Errorf("%d/%d: %w", i, int(xls.row), err)
		}
		isNil := v == nil
		if !isNil {
			if vr, ok := v.(driver.Valuer); ok {
				if vv, err := vr.Value(); err == nil {
					v = vv
				}
			}
			//fmt.Printf("%s: %#v (%T)\n", axis, v, v)
			var err error
			var printed bool
			switch x := v.(type) {
			case time.Time:
				if isNil = x.IsZero(); !isNil {
					err = xls.xl.SetCellStr(xls.Name, axis, x.Format("2006-01-02"))
					printed = true
				}
			case sql.NullTime:
				if x.Valid {
					t := x.Time
					if isNil = t.IsZero(); !isNil {
						err = xls.xl.SetCellStr(xls.Name, axis, t.Format("2006-01-02"))
						printed = true
					}
				} else {
					isNil = true
				}
			case sql.NullFloat64:
				if x.Valid {
					err = xls.xl.SetCellFloat(xls.Name, axis, x.Float64, -1, 64)
					printed = true
				} else {
					isNil = true
				}
			case sql.NullInt64:
				if x.Valid {
					err = xls.xl.SetCellInt(xls.Name, axis, int(x.Int64))
					printed = true
				} else {
					isNil = true
				}
			case sql.NullString:
				if x.Valid {
					v = x.String
				} else {
					v, isNil = "", true
				}
			case fmt.Stringer:
				v = x.String()
			}
			if isNil {
				continue
			}
			if err != nil {
				return fmt.Errorf("%s[%s]: %w", xls.Name, axis, err)
			}
			if printed {
				continue
			}
			if s, ok := v.(string); ok {
				err = xls.xl.SetCellStr(xls.Name, axis, s)
			} else {
				err = xls.xl.SetCellValue(xls.Name, axis, v)
			}
			if err != nil {
				return fmt.Errorf("%s[%s]: %w", xls.Name, axis, err)
			}

		}
	}
	return nil
}
