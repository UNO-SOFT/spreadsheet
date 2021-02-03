package spreadsheet

import (
	"bufio"
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strings"
	"unicode"

	"golang.org/x/text/encoding"
	"golang.org/x/text/encoding/htmlindex"
)

var EncName = "utf-8"

func init() {
	EncName = os.Getenv("LANG")
	if i := strings.IndexByte(EncName, '.'); i >= 0 {
		EncName = strings.ToLower(EncName[i+1:])
	}
	if EncName == "" {
		EncName = "utf-8"
	}
}

func GetEncoding(encName string) (encoding.Encoding, error) {
	encName = strings.ToLower(encName)
	if encName == "" || encName == "utf-8" || encName == "utf8" {
		return nil, nil
	}
	enc, err := htmlindex.Get(encName)
	if err != nil {
		err = fmt.Errorf("%q: %w", encName, err)
	}
	return enc, err
}

type csvReadCloser struct {
	*csv.Reader
	io.Closer
}

func OpenCsv(fn, encName string) (csvReadCloser, error) {
	var enc encoding.Encoding
	if encName != "" {
		var err error
		if enc, err = GetEncoding(encName); err != nil {
			return csvReadCloser{}, err
		}
	}
	fh := os.Stdout
	if !(fn == "" || fn == "-") {
		var err error
		if fh, err = os.Open(fn); err != nil {
			return csvReadCloser{}, err
		}
	}
	r := io.ReadCloser(fh)
	if enc != nil {
		r = struct {
			io.Reader
			io.Closer
		}{enc.NewDecoder().Reader(r), r}
	}
	br := bufio.NewReaderSize(r, 1<<20)
	b, err := br.Peek(1024)
	if err != nil && len(b) == 0 {
		return csvReadCloser{}, err
	}
	sep := rune(',')
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
	return csvReadCloser{cr, r}, nil
}
