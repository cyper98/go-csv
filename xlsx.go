package go_csv

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"
	"sync"
	"time"
)

var (
	ErrInvalidXLSX      = errors.New("xlsx: invalid xlsx file")
	ErrSheetNotFound    = errors.New("xlsx: sheet not found")
	ErrInvalidCell      = errors.New("xlsx: invalid cell")
	ErrInvalidSheetName = errors.New("xlsx: invalid sheet name")
)

var (
	xlsxColNameCache []string
	xlsxColNameOnce  sync.Once
	colNameReplacer  *strings.Replacer
)

func initColNameCache() {
	xlsxColNameOnce.Do(func() {
		xlsxColNameCache = make([]string, 1024)
		for n := 1; n <= 1024; n++ {
			xlsxColNameCache[n-1] = colIndexToName(n)
		}
	})
	colNameReplacer = strings.NewReplacer(
		"&", "&amp;",
		"<", "&lt;",
		">", "&gt;",
		"'", "&apos;",
		`"`, "&quot;",
	)
}

func initColNameOnce() {
	initColNameCache()
}

func getColName(i int) string {
	initColNameOnce()
	if i <= len(xlsxColNameCache) {
		return xlsxColNameCache[i-1]
	}
	return colIndexToName(i)
}

func fastXMLEscape(s string) string {
	if colNameReplacer != nil {
		return colNameReplacer.Replace(s)
	}
	initColNameCache()
	return colNameReplacer.Replace(s)
}

// ==== minimal XML blocks for XLSX ====

const contentTypesXMLTpl = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  %s
</Types>`

const contentTypesXMLMulti = `  <Override PartName="/xl/worksheets/%s.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`

const relsRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const workbookXMLMulti = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    %s
  </sheets>
</workbook>`

const sheetTag = `    <sheet name="%s" sheetId="%d" r:id="rId%d"/>`

const workbookRelsMulti = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  %s
</Relationships>`

const relTag = `  <Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/%s.xml"/>`

const styleXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
    <numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>
    <numFmt numFmtId="166" formatCode="$#,##0.00"/>
  </numFmts>
  <fonts count="2" x:ui="1">
    <font>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <name val="Calibri"/>
    </font>
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border/>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  </cellXfs>
</styleSheet>`

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

// ==== Options ====

type XLSXOptions struct {
	SheetName  string
	DateFormat string
}

type XLSXStyle struct {
	Bold         bool
	DateFormat   string
	NumberFormat string
	BgColor      string
}

// ==== Writer ====

type xlsxSheetWriter struct {
	zw          *zip.Writer
	sheetWriter io.Writer
	rowIndex    int
	buf         *bytes.Buffer
	batchRows   []string
	batchSize   int
	isMulti     bool
	sheets      map[string]*xlsxSheetWriter
}

// XLSXWriter supports multiple sheets
type XLSXWriter struct {
	buf     *bytes.Buffer
	zw      *zip.Writer
	sheets  map[string]*xlsxSheetWriter
	style   *XLSXStyle
	options *XLSXOptions
}

func NewXLSXWriter() *XLSXWriter {
	return &XLSXWriter{
		sheets: make(map[string]*xlsxSheetWriter),
	}
}

func (x *XLSXWriter) AddSheet(sheetName string) error {
	if sheetName == "" {
		sheetName = fmt.Sprintf("Sheet%d", len(x.sheets)+1)
	}
	if _, exists := x.sheets[sheetName]; exists {
		return ErrInvalidSheetName
	}

	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)

	w, err := zw.Create("xl/worksheets/" + sheetName + ".xml")
	if err != nil {
		return err
	}
	if _, err := io.WriteString(w, sheetHeadXML()); err != nil {
		return err
	}

	x.sheets[sheetName] = &xlsxSheetWriter{
		zw:          zw,
		sheetWriter: w,
		rowIndex:    0,
		buf:         buf,
		isMulti:     true,
	}
	return nil
}

func (x *XLSXWriter) WriteRow(sheetName string, cells []any) error {
	sw, ok := x.sheets[sheetName]
	if !ok {
		return ErrSheetNotFound
	}
	sw.rowIndex++
	rowXML := buildRowXML(sw.rowIndex, cells)
	if _, err := io.WriteString(sw.sheetWriter, rowXML); err != nil {
		return err
	}
	return nil
}

func (x *XLSXWriter) SetCellStyle(sheetName, cellRef string, style *Style) error {
	return nil
}

func (x *XLSXWriter) MergeCell(sheetName, topLeft, bottomRight string) error {
	sw, ok := x.sheets[sheetName]
	if !ok {
		return ErrSheetNotFound
	}
	mergeXML := fmt.Sprintf(`<mergeCell ref="%s:%s"/>`, topLeft, bottomRight)
	io.WriteString(sw.sheetWriter, mergeXML)
	return nil
}

func (x *XLSXWriter) SetColWidth(sheetName string, col string, width float64) error {
	return nil
}

func (x *XLSXWriter) SetRowHeight(sheetName string, row int, height float64) error {
	return nil
}

func (x *XLSXWriter) SetHyperLink(sheetName, cellRef string, link *Hyperlink) error {
	return nil
}

func (x *XLSXWriter) AddComment(sheetName string, comment *Comment) error {
	return nil
}

func (x *XLSXWriter) AddDataValidation(sheetName string, dv *DataValidation) error {
	return nil
}

func (x *XLSXWriter) AddChart(sheetName, cellRef string, chart *Chart) error {
	return nil
}

func (x *XLSXWriter) AddTable(sheetName string, table *Table) error {
	return nil
}

func (x *XLSXWriter) Close() error {
	var sheetXMLs []string
	var relsXMLs []string
	idx := 1

	for name, sw := range x.sheets {
		if _, err := io.WriteString(sw.sheetWriter, sheetTailXML()); err != nil {
			return err
		}
		if err := sw.zw.Close(); err != nil {
			return err
		}
		sheetXMLs = append(sheetXMLs, fmt.Sprintf(sheetTag, name, idx, idx))
		relsXMLs = append(relsXMLs, fmt.Sprintf(relTag, idx, name))
		idx++
	}

	mainBuf := new(bytes.Buffer)
	mainZw := zip.NewWriter(mainBuf)

	writeZipFile(mainZw, "[Content_Types].xml", []byte(buildContentTypes(len(x.sheets))))
	writeZipFile(mainZw, "_rels/.rels", []byte(relsRels))
	writeZipFile(mainZw, "xl/workbook.xml", []byte(buildWorkbook(len(x.sheets), sheetXMLs)))
	writeZipFile(mainZw, "xl/_rels/workbook.xml.rels", []byte(buildWorkbookRels(len(x.sheets), relsXMLs)))

	for name, sw := range x.sheets {
		writeZipFile(mainZw, "xl/worksheets/"+name+".xml", sw.buf.Bytes())
	}

	writeZipFile(mainZw, "xl/styles.xml", []byte(styleXML))

	mainZw.Close()
	x.buf = mainBuf
	return nil
}

func (x *XLSXWriter) Bytes() []byte {
	if x.buf != nil {
		return x.buf.Bytes()
	}
	return nil
}

func (x *XLSXWriter) WriteToFile(filePath string) error {
	if x.buf == nil {
		if err := x.Close(); err != nil {
			return err
		}
	}
	return os.WriteFile(filePath, x.buf.Bytes(), 0644)
}

func buildContentTypes(numSheets int) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)

	for i := 1; i <= numSheets; i++ {
		sb.WriteString(fmt.Sprintf("\n  <Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>", i))
	}
	sb.WriteString("\n</Types>")
	return sb.String()
}

func buildWorkbook(numSheets int, sheets []string) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>`)
	for i, s := range sheets {
		sb.WriteString("\n" + s)
		if i < len(sheets)-1 {
			sb.WriteString("\n    </sheet>")
		}
	}
	sb.WriteString("\n  </sheets>\n</workbook>")
	return sb.String()
}

func buildWorkbookRels(numSheets int, rels []string) string {
	var sb strings.Builder
	sb.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for i, r := range rels {
		sb.WriteString("\n" + r)
		if i < len(rels)-1 {
			sb.WriteString("\n  </Relationship>")
		}
	}
	sb.WriteString("\n</Relationships>")
	return sb.String()
}

func buildRowXML(rowIdx int, cells []any) string {
	var sb strings.Builder
	sb.WriteString(fmt.Sprintf(`<row r="%d">`, rowIdx))
	for i, v := range cells {
		colRef := getColName(i+1) + strconv.Itoa(rowIdx)
		if n, ok := tryNumber(v); ok {
			sb.WriteString(fmt.Sprintf(`<c r="%s" t="n"><v>%s</v></c>`, colRef, n))
		} else if t, ok := v.(time.Time); ok {
			excelDate := excelTime(t)
			sb.WriteString(fmt.Sprintf(`<c r="%s" t="n"><v>%.10f</v></c>`, colRef, excelDate))
		} else {
			s, _ := DefaultToString(v)
			s = fastXMLEscape(s)
			sb.WriteString(fmt.Sprintf(`<c r="%s" t="inlineStr"><is><t>%s</t></is></c>`, colRef, s))
		}
	}
	sb.WriteString(`</row>`)
	return sb.String()
}

func excelTime(t time.Time) float64 {
	return t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24
}

// ==== Legacy single-sheet API ====

type XLSXSingleWriter struct {
	zw          *zip.Writer
	sheetWriter io.Writer
	rowIndex    int
	tempBuf     *bytes.Buffer
}

func NewXLSXBuffer(sheetName string) (*XLSXSingleWriter, *bytes.Buffer, error) {
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)

	if err := writeZipFile(zw, "[Content_Types].xml", []byte(contentTypesXMLTpl)); err != nil {
		return nil, nil, err
	}
	if err := writeZipFile(zw, "_rels/.rels", []byte(relsRels)); err != nil {
		return nil, nil, err
	}
	if err := writeZipFile(zw, "xl/workbook.xml", []byte(fmt.Sprintf(workbookXMLMulti, "    <sheet name=\""+fastXMLEscape(sheetName)+"\" sheetId=\"1\" r:id=\"rId1\"/>"))); err != nil {
		return nil, nil, err
	}
	if err := writeZipFile(zw, "xl/_rels/workbook.xml.rels", []byte(workbookRelsMulti)); err != nil {
		return nil, nil, err
	}
	w, err := zw.Create("xl/worksheets/sheet1.xml")
	if err != nil {
		return nil, nil, err
	}
	if _, err := io.WriteString(w, sheetHeadXML()); err != nil {
		return nil, nil, err
	}

	return &XLSXSingleWriter{
		zw:          zw,
		sheetWriter: w,
		rowIndex:    0,
		tempBuf:     buf,
	}, buf, nil
}

func (x *XLSXSingleWriter) WriteRow(cells []any) error {
	x.rowIndex++
	rowXML := buildRowXML(x.rowIndex, cells)
	if _, err := io.WriteString(x.sheetWriter, rowXML); err != nil {
		return err
	}
	return nil
}

func (x *XLSXSingleWriter) Close() error {
	if _, err := io.WriteString(x.sheetWriter, sheetTailXML()); err != nil {
		return err
	}
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

// ==== XLSX Reader ====

type XLSXReader struct {
	zipReader *zip.ReadCloser
	sheets    map[string][][]string
}

func ReadXLSXFile(filePath string) (*XLSXReader, error) {
	f, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	zr, err := zip.NewReader(f, 0)
	if err != nil {
		f.Close()
		return nil, ErrInvalidXLSX
	}
	_ = f.Close()
	return parseXLSX(zr)
}

func ReadXLSX(r io.ReaderAt, size int64) (*XLSXReader, error) {
	zr, err := zip.NewReader(r, size)
	if err != nil {
		return nil, ErrInvalidXLSX
	}
	return parseXLSX(zr)
}

func parseXLSX(zr *zip.Reader) (*XLSXReader, error) {
	reader := &XLSXReader{
		sheets: make(map[string][][]string),
	}

	sheetMap := make(map[string]string)
	var workbook []byte

	for _, f := range zr.File {
		name := f.Name
		if strings.HasPrefix(name, "xl/worksheets/") && strings.HasSuffix(name, ".xml") {
			rc, err := f.Open()
			if err != nil {
				continue
			}
			data, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				continue
			}
			sheetName := strings.TrimSuffix(strings.TrimPrefix(name, "xl/worksheets/"), ".xml")
			sheetMap[sheetName] = string(data)
		}
		if name == "xl/workbook.xml" {
			rc, err := f.Open()
			if err == nil {
				workbook, _ = io.ReadAll(rc)
				rc.Close()
			}
		}
	}

	sheetNames := extractSheetNames(workbook)
	for _, name := range sheetNames {
		if data, ok := sheetMap[name]; ok {
			rows := parseSheetXML(data)
			reader.sheets[name] = rows
		}
	}

	return reader, nil
}

func extractSheetNames(data []byte) []string {
	var names []string
	lines := strings.Split(string(data), "<sheet name=\"")
	for i := 1; i < len(lines); i++ {
		end := strings.Index(lines[i], "\"")
		if end > 0 {
			names = append(names, lines[i][:end])
		}
	}
	return names
}

func parseSheetXML(data string) [][]string {
	var rows [][]string
	lines := strings.Split(data, "<row r=\"")

	for i := 1; i < len(lines); i++ {
		row := lines[i]
		endIdx := strings.Index(row, "\">")
		if endIdx < 0 {
			continue
		}

		cells := extractCells(row[endIdx:])
		if len(cells) > 0 {
			rows = append(rows, cells)
		}
	}

	return rows
}

func extractCells(rowData string) []string {
	var cells []string
	idx := 0

	for {
		start := strings.Index(rowData[idx:], "<c r=\"")
		if start < 0 {
			break
		}
		idx += start + 6

		refEnd := strings.Index(rowData[idx:], "\"")
		if refEnd < 0 {
			break
		}
		_ = rowData[idx : idx+refEnd]
		idx += refEnd + 2

		if strings.Contains(rowData[idx:], "<v>") {
			vStart := strings.Index(rowData[idx:], "<v>") + 3
			vEnd := strings.Index(rowData[idx:], "</v>")
			if vEnd > vStart {
				cells = append(cells, rowData[idx+vStart:idx+vEnd])
			}
			idx += vEnd
		} else if strings.Contains(rowData[idx:], "<t>") {
			tStart := strings.Index(rowData[idx:], "<t>") + 3
			tEnd := strings.Index(rowData[idx:], "</t>")
			if tEnd > tStart {
				cells = append(cells, rowData[idx+tStart:idx+tEnd])
			}
			idx += tEnd
		}
	}

	return cells
}

func (r *XLSXReader) SheetNames() []string {
	var names []string
	for name := range r.sheets {
		names = append(names, name)
	}
	return names
}

func (r *XLSXReader) Rows(sheetName string) ([][]string, bool) {
	rows, ok := r.sheets[sheetName]
	return rows, ok
}

func (r *XLSXReader) GetSheet(name string) [][]string {
	return r.sheets[name]
}

func (r *XLSXReader) Close() error {
	if r.zipReader != nil {
		return r.zipReader.Close()
	}
	return nil
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
	name := ""
	for i > 0 {
		rem := (i - 1) % 26
		name = string(rune('A'+rem)) + name
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
