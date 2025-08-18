package go_csv

import (
	"bufio"
	"encoding/csv"
	"io"
	"os"
)

type CSVOptions struct {
	Comma   rune // default ','
	UseCRLF bool // \r\n
	BOM     bool // write UTF-8 BOM
}

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
