# go-csv

Export **CSV**, **XLSX**, and bundle multiple files into **ZIP** in a style similar to “fast-csv” (Node.js) but written in **Go**.  
Designed for **streaming**, each row is a `[]any`, auto-converted to string/number, with support for headers, BOM, CRLF, and ZIP compression.

- **No external libraries**: only `encoding/csv`, `archive/zip`, `encoding/xml`, etc.
- **Streaming support**: write rows one by one, avoiding memory blow-up.
- **Simple API**: just pass `[]any` per row, like fast-csv.
- **ZIP support**: compress from in-memory buffers or existing file paths.

## Requirements
- Go 1.20+ (recommended 1.21+)

## Installation

```bash
go get github.com/aminofox/go-csv
```

Project structure:
```
go-csv/
  csv.go
  xlsx.go
  zip.go
  util.go
  examples/
    main.go
  README.md
  go.mod
```

## Main API

### CSV
```go
// Write to file
func WriteCSVFile(
  filePath string,
  headers []any,
  data [][]any,
  opt *CSVOptions,
  toString ToStringFunc,
) error

// Write to writer (streaming with rows <-chan []any)
func WriteCSVToWriter(
  w io.Writer,
  headers []any,
  rows <-chan []any,
  opt *CSVOptions,
  toString ToStringFunc,
) error

type CSVOptions struct {
  Comma  rune  // default ','
  UseCRLF bool // \r\n
  BOM    bool  // write UTF-8 BOM
}

// Custom type-to-string converter
type ToStringFunc func(v any) (string, bool)
```

### XLSX
Minimal XLSX (no external libraries), built using `archive/zip` + XML. Text cells use `inlineStr`, numbers use `t="n"`.

```go
// Write to file (single sheet)
func WriteXLSXFile(
  filePath string,
  sheetName string,
  headers []any,
  data [][]any,
) error

// Write to writer (streaming with rows <-chan []any)
func WriteXLSXToWriter(
  w io.Writer,
  sheetName string,
  headers []any,
  rows <-chan []any,
) error
```

### ZIP
```go
// Compress in-memory buffers into a zip
func ZipBuffers(outputZipPath string, files map[string][]byte) error

// Compress existing files on disk
func ZipPaths(outputZipPath string, paths []string, baseDir string) error
```

## Quick Usage

### 1) Write CSV from `[][]any`
```go
headers := []any{"id", "name", "score", "active"}
data := [][]any{
  {1, "alice", 9.5, true},
  {2, "bob", 8.1, false},
  {3, "seang", 10, true},
}

err := export.WriteCSVFile(
  "out/users.csv",
  headers,
  data,
  &export.CSVOptions{Comma: ',', BOM: true},
  nil, // no custom toString
)
if err != nil { panic(err) }
```

### 2) Streaming CSV (row by row, avoids loading all into memory)
```go
f, _ := os.Create("out/users_stream.csv")
defer f.Close()

rows := make(chan []any, 1024)
go func() {
  for _, r := range data { rows <- r }
  close(rows)
}()

err := export.WriteCSVToWriter(
  f,
  headers,
  rows,
  &export.CSVOptions{BOM: true},
  nil,
)
if err != nil { panic(err) }
```

### 3) Write XLSX from `[][]any`
```go
err := export.WriteXLSXFile("out/users.xlsx", "Users", headers, data)
if err != nil { panic(err) }
```

### 4) Streaming XLSX
```go
var buf bytes.Buffer
rowch := make(chan []any, 4)
go func() {
  rowch <- []any{"A", "B", "C"}
  rowch <- []any{1, 2, 3}
  close(rowch)
}()

if err := export.WriteXLSXToWriter(&buf, "StreamSheet", nil, rowch); err != nil {
  panic(err)
}
_ = os.WriteFile("out/stream.xlsx", buf.Bytes(), 0644)
```

### 5) ZIP multiple files
- **From buffers**:
```go
csvBytes, _ := os.ReadFile("out/users.csv")
xlsxBytes, _ := os.ReadFile("out/users.xlsx")

files := map[string][]byte{
  "csv/users.csv":   csvBytes,
  "xlsx/users.xlsx": xlsxBytes,
}
if err := export.ZipBuffers("out/exports.zip", files); err != nil {
  panic(err)
}
```

- **From file paths**:
```go
paths := []string{"out/users.csv", "out/users_stream.csv", "out/users.xlsx"}
if err := export.ZipPaths("out/exports_paths.zip", paths, "out"); err != nil {
  panic(err)
}
```

## Custom toString
If you want to control formatting (e.g. currency, datetime), provide a `toString`:
```go
customToString := func(v any) (string, bool) {
  switch t := v.(type) {
  case float64:
    return fmt.Sprintf("%.2f", t), true
  case time.Time:
    return t.In(time.FixedZone("UTC+7", 7*3600)).Format("2006-01-02 15:04:05"), true
  default:
    return "", false // fallback to DefaultToString
  }
}

_ = export.WriteCSVFile("out/custom.csv", headers, data, &export.CSVOptions{BOM:true}, customToString)
```

## Notes & Limitations
- **XLSX**:
  - Single sheet (`sheet1.xml`) for now.
  - Text as `inlineStr`, numbers as `t="n"`. No styles/formatting (dates, number formats).
- **CSV**: set `Comma` for different locales; enable `BOM: true` for Excel on Windows.
- **Streaming**: use `chan []any` to handle very large datasets.
- **Zero dependencies**: pure Go stdlib.

## Performance Tips
- Use **streaming** for large datasets (millions of rows).
- Increase channel buffer (`rows := make(chan []any, 4096)`).
- Minimize string formatting in hot paths; centralize in `toString`.

## Full Example
See `examples/main.go` in the repo:

```bash
go run ./examples
```

## License
MIT
