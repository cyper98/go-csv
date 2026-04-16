package go_csv

import (
	"bufio"
	"bytes"
	"encoding/csv"
	"errors"
	"io"
	"os"
	"sync"
)

var (
	ErrEmptyHeader   = errors.New("csv: empty header required")
	ErrInvalidRow    = errors.New("csv: invalid row")
	ErrMismatchedLen = errors.New("csv: row length mismatch")
)

var stringSlicePool = sync.Pool{
	New: func() interface{} { return make([]string, 0, 32) },
}

func getStringSlice() []string {
	return stringSlicePool.Get().([]string)[:0]
}

func putStringSlice(s []string) {
	s = s[:0]
	stringSlicePool.Put(s)
}

type CSVOptions struct {
	Comma    rune // default ','
	UseCRLF  bool // \r\n
	BOM      bool // write UTF-8 BOM
	Comment  rune // comment character (default '#')
	QuoteAll bool // always quote fields
}

type CSVReadOptions struct {
	Comma     rune // default ','
	HasHeader bool // first row is header
	SkipEmpty bool // skip empty rows
	TrimSpace bool // trim leading/trailing space
}

type ValidationFunc func(rowIndex int, row []any) error

// convert any value to string (based on util.go)
type ToStringFunc func(v any) (string, bool)

func WriteCSVToWriter(w io.Writer, headers []any, rows <-chan []any, opt *CSVOptions, toString ToStringFunc) error {
	if opt == nil {
		opt = &CSVOptions{}
	}
	bw := bufio.NewWriter(w)
	defer bw.Flush()

	// BOM
	if opt.BOM {
		if _, err := bw.Write([]byte{0xEF, 0xBB, 0xBF}); err != nil {
			return err
		}
	}

	cw := csv.NewWriter(bw)
	if opt.Comma != 0 {
		cw.Comma = opt.Comma
	}
	cw.UseCRLF = opt.UseCRLF

	// headers
	if len(headers) > 0 {
		record := make([]string, len(headers))
		for i, h := range headers {
			record[i], _ = DefaultToString(h)
			if toString != nil {
				if s, ok := toString(h); ok {
					record[i] = s
				}
			}
		}
		if err := cw.Write(record); err != nil {
			return err
		}
	}

	// rows streaming
	for row := range rows {
		record := make([]string, len(row))
		for i, v := range row {
			record[i], _ = DefaultToString(v)
			if toString != nil {
				if s, ok := toString(v); ok {
					record[i] = s
				}
			}
		}
		if err := cw.Write(record); err != nil {
			return err
		}
	}

	cw.Flush()
	return cw.Error()
}

func WriteCSVFile(filePath string, headers []any, data [][]any, opt *CSVOptions, toString ToStringFunc) error {
	f, err := os.Create(filePath)
	if err != nil {
		return err
	}
	defer f.Close()

	ch := make(chan []any, 1024)
	go func() {
		for _, r := range data {
			ch <- r
		}
		close(ch)
	}()

	return WriteCSVToWriter(f, headers, ch, opt, toString)
}

func ReadCSVFile(filePath string, opt *CSVReadOptions) (headers []string, rows <-chan []string, err error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, nil, err
	}
	return ReadCSVFromReader(f, opt)
}

func ReadCSVFromReader(r io.Reader, opt *CSVReadOptions) (headers []string, rows <-chan []string, err error) {
	if opt == nil {
		opt = &CSVReadOptions{}
	}
	if opt.Comma == 0 {
		opt.Comma = ','
	}

	br := bufio.NewReader(r)
	bom := []byte{0xEF, 0xBB, 0xBF}
	buf := make([]byte, 3)
	n, _ := br.Read(buf)
	if n < 3 || string(buf) != string(bom) {
		br.Reset(io.NopCloser(bytes.NewReader(buf[:n])))
	}

	cr := csv.NewReader(br)
	cr.Comma = opt.Comma
	cr.TrimLeadingSpace = opt.TrimSpace
	cr.ReuseRecord = false

	headerRow := []string{}
	ch := make(chan []string, 512)

	for {
		record, err := cr.Read()
		if err != nil {
			if errors.Is(err, io.EOF) {
				break
			}
			return nil, nil, err
		}
		if opt.SkipEmpty && len(record) == 0 {
			continue
		}
		if len(headerRow) == 0 && opt.HasHeader {
			headerRow = make([]string, len(record))
			copy(headerRow, record)
			continue
		}
		ch <- record
	}
	close(ch)

	return headerRow, ch, nil
}

func ReadAllCSV(r io.Reader, opt *CSVReadOptions) (headers []string, rows [][]string, err error) {
	h, ch, err := ReadCSVFromReader(r, opt)
	if err != nil {
		return nil, nil, err
	}
	for row := range ch {
		rows = append(rows, row)
	}
	return h, rows, nil
}

func ReadCSVAsAny(r io.Reader, opt *CSVReadOptions) (headers []string, rows <-chan []any, err error) {
	_, strCh, err := ReadCSVFromReader(r, opt)
	if err != nil {
		return nil, nil, err
	}
	ch := make(chan []any, 512)

	go func(h []string) {
		defer close(ch)
		for row := range strCh {
			anyRow := make([]any, len(row))
			for i, v := range row {
				anyRow[i] = v
			}
			ch <- anyRow
		}
	}(headers)

	return headers, ch, nil
}
