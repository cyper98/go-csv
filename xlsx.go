package go_csv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"strconv"
)

// ==== minimal XML blocks for XLSX ====

const contentTypesXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

const relsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const workbookXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="%s" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`

const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

// worksheet head/tail, sheet data will record stream between head/tail
func sheetHeadXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>`
}

func sheetTailXML() string {
	return `  </sheetData>
</worksheet>`
}

// ==== Writer ====

type XLSXOptions struct {
	SheetName string // default "Sheet1"
}

type xlsxSheetWriter struct {
	zw          *zip.Writer
	sheetWriter io.Writer
	rowIndex    int
}

func NewXLSXBuffer(sheetName string) (*xlsxSheetWriter, *bytes.Buffer, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)

	// [Content_Types].xml
	if err := writeZipFile(zw, "[Content_Types].xml", []byte(contentTypesXML)); err != nil {
		return nil, nil, err
	}
	// _rels/.rels
	if err := writeZipFile(zw, "_rels/.rels", []byte(relsRels)); err != nil {
		return nil, nil, err
	}
	// xl/workbook.xml
	if err := writeZipFile(zw, "xl/workbook.xml", []byte(fmt.Sprintf(workbookXML, xmlEscape(sheetName)))); err != nil {
		return nil, nil, err
	}
	// xl/_rels/workbook.xml.rels
	if err := writeZipFile(zw, "xl/_rels/workbook.xml.rels", []byte(workbookRels)); err != nil {
		return nil, nil, err
	}
	// xl/worksheets/sheet1.xml (stream)
	w, err := zw.Create("xl/worksheets/sheet1.xml")
	if err != nil {
		return nil, nil, err
	}
	if _, err := io.WriteString(w, sheetHeadXML()); err != nil {
		return nil, nil, err
	}

	return &xlsxSheetWriter{
		zw:          zw,
		sheetWriter: w,
		rowIndex:    0,
	}, buf, nil
}

func (x *xlsxSheetWriter) WriteRow(cells []any) error {
	x.rowIndex++
	// each row: <row r="1"> ... </row>
	if _, err := io.WriteString(x.sheetWriter, fmt.Sprintf(`<row r="%d">`, x.rowIndex)); err != nil {
		return err
	}
	for i, v := range cells {
		colRef := colIndexToName(i+1) + strconv.Itoa(x.rowIndex)
		// detect number or string
		if n, ok := tryNumber(v); ok {
			// numeric cell
			if _, err := io.WriteString(x.sheetWriter, fmt.Sprintf(`<c r="%s" t="n"><v>%s</v></c>`, colRef, n)); err != nil {
				return err
			}
		} else {
			// inlineStr cell
			s, _ := DefaultToString(v)
			s = xmlEscape(s)
			if _, err := io.WriteString(x.sheetWriter, fmt.Sprintf(`<c r="%s" t="inlineStr"><is><t>%s</t></is></c>`, colRef, s)); err != nil {
				return err
			}
		}
	}
	if _, err := io.WriteString(x.sheetWriter, `</row>`); err != nil {
		return err
	}
	return nil
}

func (x *xlsxSheetWriter) Close() error {
	// tail
	if _, err := io.WriteString(x.sheetWriter, sheetTailXML()); err != nil {
		return err
	}
	// close zip
	return x.zw.Close()
}

func WriteXLSXFile(filePath string, sheetName string, headers []any, data [][]any) error {
	xw, buf, err := NewXLSXBuffer(sheetName)
	if err != nil {
		return err
	}

	if len(headers) > 0 {
		if err := xw.WriteRow(headers); err != nil {
			return err
		}
	}
	for _, r := range data {
		if err := xw.WriteRow(r); err != nil {
			return err
		}
	}
	if err := xw.Close(); err != nil {
		return err
	}

	return os.WriteFile(filePath, buf.Bytes(), 0644)
}

func WriteXLSXToWriter(w io.Writer, sheetName string, headers []any, rows <-chan []any) error {
	xw, buf, err := NewXLSXBuffer(sheetName)
	if err != nil {
		return err
	}

	if len(headers) > 0 {
		if err := xw.WriteRow(headers); err != nil {
			return err
		}
	}
	for r := range rows {
		if err := xw.WriteRow(r); err != nil {
			return err
		}
	}
	if err := xw.Close(); err != nil {
		return err
	}
	_, err = w.Write(buf.Bytes())
	return err
}

// ==== helpers ====

func writeZipFile(zw *zip.Writer, name string, data []byte) error {
	w, err := zw.Create(name)
	if err != nil {
		return err
	}
	_, err = w.Write(data)
	return err
}

func colIndexToName(i int) string {
	// 1 -> A, 26-> Z, 27 -> AA
	name := ""
	for i > 0 {
		rem := (i - 1) % 26
		name = string('A'+rem) + name
		i = (i - 1) / 26
	}
	return name
}

func tryNumber(v any) (string, bool) {
	switch t := v.(type) {
	case int, int32, int64:
		return fmt.Sprintf("%v", t), true
	case uint, uint32, uint64:
		return fmt.Sprintf("%v", t), true
	case float32:
		return strconv.FormatFloat(float64(t), 'f', -1, 64), true
	case float64:
		return strconv.FormatFloat(t, 'f', -1, 64), true
	default:
		return "", false
	}
}

func xmlEscape(s string) string {
	var b bytes.Buffer
	xml.EscapeText(&b, []byte(s))
	return b.String()
}
