// Copyright 2020, Tamás Gulácsi.
//
// SPDX-License-Identifier: Apache-2.0

package spreadsheet

import "io"

// Writer writes the spreadsheet consisting of the sheets created
// with NewSheet. The write finishes when Close is called.
//
// The writer SHOULD allow writing to separate sheets concurrently,
// and document if it does not provide this functionality.
type Writer interface {
	io.Closer
	NewSheet(name string, cols []Column) (Sheet, error)
}

// Sheet should be Closed when finished.
type Sheet interface {
	io.Closer
	AppendRow(values ...interface{}) error
}

// Style is a style for a column/row/cell.
type Style struct {
	// FontBold is true if the font is bold
	FontBold bool
	// Format is the number format
	Format string
}

// Column contains the Name of the column and header's style and column's style.
type Column struct {
	Name           string
	Header, Column Style
}
