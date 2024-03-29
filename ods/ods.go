// Copyright 2020, 2023 Tamás Gulácsi.
//
// SPDX-License-Identifier: Apache-2.0

package ods

import (
	"encoding/xml"
	"fmt"
	"hash/fnv"
	"io"
	"os"
	"strings"
	"sync"
	"time"

	"github.com/klauspost/compress/zip"
	"github.com/klauspost/compress/zstd"
	qt "github.com/valyala/quicktemplate"

	"github.com/UNO-SOFT/spreadsheet"
)

var _ = fmt.Errorf

//go:generate qtc

var qtMu sync.Mutex

// acquireWriter wraps the given io.Writer to be usable with quicktemplates.
func acquireWriter(w io.Writer) *qt.Writer {
	qtMu.Lock()
	W := qt.AcquireWriter(w)
	qtMu.Unlock()
	return W
}

// releaseWriter returns the *quicktemplate.Writer to the pool.
func releaseWriter(W *qt.Writer) { qtMu.Lock(); qt.ReleaseWriter(W); qtMu.Unlock() }

// ValueType is the cell's value's type.
type ValueType struct {
	t uint8
}

func (v ValueType) String() string {
	switch v.t {
	case 'f':
		return "float"
	case 'd':
		return "date"
	default:
		return "string"
	}
}
func getValueType(v interface{}) ValueType {
	switch x := v.(type) {
	case float32, float64,
		int, int8, int16, int32, int64,
		uint, uint16, uint32, uint64:
		return FloatType
	case time.Time:
		return DateType
	case string:
		if strings.HasPrefix(x, "https://") || strings.HasPrefix(x, "http://") {
			return LinkType
		}
		return StringType
	default:
		return StringType
	}
}

var (
	// FloatType for numerical data
	FloatType = ValueType{'f'}
	// DateType for dates
	DateType = ValueType{'d'}
	// LinkType is a string that seems to be a https?:// link
	LinkType = ValueType{'a'}
	// StringType for everything else
	StringType = ValueType{'s'}
)

// NewWriter returns a content writer and a zip closer for an ods file.
//
// This writer allows concurrent write to separate sheets.
func NewWriter(w io.Writer) (*ODSWriter, error) {
	now := time.Now()
	zw := zip.NewWriter(w)
	for _, elt := range []struct {
		Stream func(*qt.Writer)
		Name   string
	}{
		{Name: "mimetype", Stream: StreamMimetype},
		{Name: "meta.xml", Stream: StreamMeta},
		{Name: "META-INF/manifest.xml", Stream: StreamManifest},
		{Name: "settings.xml", Stream: StreamSettings},
	} {
		parts := strings.SplitAfter(elt.Name, "/")
		var prev string
		for _, p := range parts[:len(parts)-1] {
			prev += p
			if _, err := zw.CreateHeader(&zip.FileHeader{
				Name: prev, Method: zip.Deflate, Modified: now,
			}); err != nil {
				return nil, err
			}
		}
		sub, err := zw.CreateHeader(&zip.FileHeader{
			Name: elt.Name, Method: zip.Deflate, Modified: now,
		})
		if err != nil {
			zw.Close()
			return nil, err
		}
		W := acquireWriter(sub)
		elt.Stream(W)
		releaseWriter(W)
	}

	bw, err := zw.CreateHeader(&zip.FileHeader{
		Name: "content.xml", Method: zip.Deflate, Modified: now,
	})
	if err != nil {
		zw.Close()
		return nil, err
	}
	W := acquireWriter(bw)
	StreamBeginSpreadsheet(W)
	releaseWriter(W)

	return &ODSWriter{w: bw, zipWriter: zw}, nil
}

// ODSWriter writes content.xml of ODS zip.
type ODSWriter struct {
	w         io.Writer
	zipWriter *zip.Writer
	styles    map[string]string
	files     []<-chan io.ReadCloser
	mu        sync.Mutex
}

// Close the ODSWriter.
func (ow *ODSWriter) Close() error {
	if ow == nil {
		return nil
	}
	ow.mu.Lock()
	defer ow.mu.Unlock()
	if ow.w == nil {
		return nil
	}

	if err := ow.copyFiles(true); err != nil {
		return err
	}
	ow.files = nil

	W := acquireWriter(ow.w)
	StreamEndSpreadsheet(W)
	releaseWriter(W)
	ow.w = nil
	zw := ow.zipWriter
	ow.zipWriter = nil
	defer zw.Close()
	bw, err := zw.CreateHeader(&zip.FileHeader{
		Name: "styles.xml", Method: zip.Deflate, Modified: time.Now(),
	})
	if err != nil {
		return err
	}
	W = acquireWriter(bw)
	StreamStyles(W, ow.styles)
	releaseWriter(W)
	return zw.Close()
}

// copyFiles copies the finished files.
func (ow *ODSWriter) copyFiles(wait bool) error {
	for _, ch := range ow.files {
		var f io.ReadCloser
		if wait {
			f = <-ch
		} else {
			select {
			case f = <-ch:
			default:
				return nil
			}
		}
		if f == nil {
			continue
		}
		_, err := io.Copy(ow.w, f)
		f.Close()
		if err != nil {
			return err
		}
	}
	return nil
}

func (ow *ODSWriter) NewSheet(name string, cols []spreadsheet.Column) (spreadsheet.Sheet, error) {
	ow.mu.Lock()
	defer ow.mu.Unlock()
	sheet := &ODSSheet{Name: name, ow: ow}

	var err error
	if sheet.f, err = os.CreateTemp("", "spreadsheet-ods-*.xml"); err != nil {
		return nil, err
	}
	os.Remove(sheet.f.Name())
	if sheet.zw, err = zstd.NewWriter(sheet.f, zstd.WithEncoderLevel(zstd.SpeedFastest)); err != nil {
		sheet.f.Close()
		return nil, err
	}
	sheet.w = acquireWriter(sheet.zw)
	ch := make(chan io.ReadCloser, 1)
	sheet.done = ch
	ow.files = append(ow.files, ch)

	ow.StreamBeginSheet(sheet.w, name, cols)
	return sheet, nil
}

func (ow *ODSWriter) getStyleName(style spreadsheet.Style) string {
	if !style.FontBold {
		return ""
	}
	hsh := fnv.New32()
	//fmt.Fprintf(hsh, "%t\t%s", style.FontBold, style.Format)
	fmt.Fprintf(hsh, "%t", style.FontBold)
	k := fmt.Sprintf("bf-%d", hsh.Sum32())
	if _, ok := ow.styles[k]; ok {
		return k
	}
	if ow.styles == nil {
		ow.styles = make(map[string]string, 1)
	}
	ow.styles[k] = `<style:style style:name="` + k + `" style:family="table-cell"><style:text-properties text:display="true" fo:font-weight="bold" /></style:style>`
	return k
}

type ODSSheet struct {
	done     chan<- io.ReadCloser
	ow       *ODSWriter
	w        *qt.Writer
	f        *os.File
	zw       *zstd.Encoder
	Name     string
	rowCount int
	mu       sync.Mutex
}

const MaxRowCount = 1 << 20

func (ods *ODSSheet) AppendRow(values ...interface{}) error {
	ods.mu.Lock()
	if ods.rowCount >= MaxRowCount {
		ods.mu.Unlock()
		return spreadsheet.ErrTooManyRows
	}
	StreamRow(ods.w, values...)
	ods.rowCount++
	ods.mu.Unlock()
	return nil
}

func (ods *ODSSheet) Close() error {
	if ods == nil {
		return nil
	}
	ods.mu.Lock()
	defer ods.mu.Unlock()

	ow, W, zw, f, done := ods.ow, ods.w, ods.zw, ods.f, ods.done
	ods.ow, ods.w, ods.zw, ods.f, ods.done = nil, nil, nil, nil, nil
	if W == nil {
		return nil
	}
	StreamEndSheet(W)
	releaseWriter(W)
	if done == nil {
		return nil
	}
	defer close(done)
	if f == nil {
		return nil
	}
	if zw != nil {
		if err := zw.Close(); err != nil {
			return err
		}
	}
	var err error
	if _, err = f.Seek(0, 0); err != nil {
		f.Close()
		return err
	}
	if zw == nil {
		done <- f
		return nil
	}
	var zr *zstd.Decoder
	if zr, err = zstd.NewReader(f); err != nil {
		return nil
	}
	done <- struct {
		io.Reader
		io.Closer
	}{
		Reader: zr,
		Closer: closerFunc(func() error {
			zr.Close()
			f.Close()
			os.Remove(f.Name())
			return nil
		}),
	}
	ow.mu.Lock()
	err = ow.copyFiles(false)
	ow.mu.Unlock()
	return err
}

// Style information - generated from content.xml with github.com/miek/zek/cmd/zek.
type Style struct {
	XMLName         xml.Name `xml:"style"`
	Name            string   `xml:"name,attr"`
	Family          string   `xml:"family,attr"`
	MasterPageName  string   `xml:"master-page-name,attr"`
	DataStyleName   string   `xml:"data-style-name,attr"`
	TableProperties struct {
		Display     string `xml:"display,attr"`
		WritingMode string `xml:"writing-mode,attr"`
	} `xml:"table-properties"`
	TextProperties struct {
		FontWeight           string `xml:"font-weight,attr"`
		FontStyle            string `xml:"font-style,attr"`
		TextPosition         string `xml:"text-position,attr"`
		TextLineThroughType  string `xml:"text-line-through-type,attr"`
		TextLineThroughStyle string `xml:"text-line-through-style,attr"`
		TextUnderlineType    string `xml:"text-underline-type,attr"`
		TextUnderlineStyle   string `xml:"text-underline-style,attr"`
		TextUnderlineWidth   string `xml:"text-underline-width,attr"`
		Display              string `xml:"display,attr"`
		TextUnderlineColor   string `xml:"text-underline-color,attr"`
		TextUnderlineMode    string `xml:"text-underline-mode,attr"`
		FontSize             string `xml:"font-size,attr"`
		Color                string `xml:"color,attr"`
		FontFamily           string `xml:"font-family,attr"`
	} `xml:"text-properties"`
	TableRowProperties struct {
		RowHeight           string `xml:"row-height,attr"`
		UseOptimalRowHeight string `xml:"use-optimal-row-height,attr"`
	} `xml:"table-row-properties"`
	TableColumnProperties struct {
		ColumnWidth           string `xml:"column-width,attr"`
		UseOptimalColumnWidth string `xml:"use-optimal-column-width,attr"`
	} `xml:"table-column-properties"`
	TableCellProperties struct {
		BackgroundColor          string `xml:"background-color,attr"`
		BorderTop                string `xml:"border-top,attr"`
		BorderBottom             string `xml:"border-bottom,attr"`
		BorderLeft               string `xml:"border-left,attr"`
		BorderRight              string `xml:"border-right,attr"`
		DiagonalBlTr             string `xml:"diagonal-bl-tr,attr"`
		DiagonalTlBr             string `xml:"diagonal-tl-br,attr"`
		VerticalAlign            string `xml:"vertical-align,attr"`
		WrapOption               string `xml:"wrap-option,attr"`
		ShrinkToFit              string `xml:"shrink-to-fit,attr"`
		WritingMode              string `xml:"writing-mode,attr"`
		GlyphOrientationVertical string `xml:"glyph-orientation-vertical,attr"`
		CellProtect              string `xml:"cell-protect,attr"`
		RotationAlign            string `xml:"rotation-align,attr"`
		RotationAngle            string `xml:"rotation-angle,attr"`
		PrintContent             string `xml:"print-content,attr"`
		DecimalPlaces            string `xml:"decimal-places,attr"`
		TextAlignSource          string `xml:"text-align-source,attr"`
		RepeatContent            string `xml:"repeat-content,attr"`
	} `xml:"table-cell-properties"`
	ParagraphProperties struct {
		WritingModeAutomatic string `xml:"writing-mode-automatic,attr"`
		MarginLeft           string `xml:"margin-left,attr"`
	} `xml:"paragraph-properties"`
}

type closerFunc func() error

func (f closerFunc) Close() error { return f() }
