package main

import (
	"bytes"
	"fmt"
	"os"

	export "github.com/aminofox/go-csv"
)

func main() {
	// ===== 1) CSV =====
	headers := []any{"id", "name", "score", "active"}
	data := [][]any{
		{1, "alice", 9.5, true},
		{2, "bob", 8.1, false},
		{3, "seang", 10, true},
	}

	if err := export.WriteCSVFile("users.csv", headers, data, &export.CSVOptions{Comma: ',', BOM: true}, nil); err != nil {
		panic(err)
	}
	fmt.Println("CSV written:", "users.csv")

	// Stream CSV
	outCSV, _ := os.Create("users_stream.csv")
	defer outCSV.Close()
	ch := make(chan []any, 4)
	go func() {
		for _, r := range data {
			ch <- r
		}
		close(ch)
	}()
	if err := export.WriteCSVToWriter(outCSV, headers, ch, &export.CSVOptions{BOM: true}, nil); err != nil {
		panic(err)
	}
	fmt.Println("CSV stream written:", "users_stream.csv")

	// ===== 2) XLSX =====
	if err := export.WriteXLSXFile("users.xlsx", "Users", headers, data); err != nil {
		panic(err)
	}
	fmt.Println("XLSX written:", "users.xlsx")

	// ===== 3) ZIP multiple files (from buffer) =====
	// example of merging CSV stream + newly created XLSX (read back into buffer)
	csvBytes, _ := os.ReadFile("users.csv")
	xlsxBytes, _ := os.ReadFile("users.xlsx")

	files := map[string][]byte{
		"csv/users.csv":   csvBytes,
		"xlsx/users.xlsx": xlsxBytes,
	}
	if err := export.ZipBuffers("exports.zip", files); err != nil {
		panic(err)
	}
	fmt.Println("ZIP written:", "exports.zip")

	// Or ZIP the available paths
	paths := []string{"users.csv", "users_stream.csv", "users.xlsx"}
	if err := export.ZipPaths("exports_paths.zip", paths, "out"); err != nil {
		panic(err)
	}
	fmt.Println("ZIP written:", "exports_paths.zip")

	// ===== 4) XLSX streams directly to writer (e.g. bytes.Buffer) =====
	var buf bytes.Buffer
	rowChan := make(chan []any, 4)
	go func() {
		rowChan <- []any{"A", "B", "C"}
		rowChan <- []any{1, 2, 3}
		close(rowChan)
	}()
	if err := export.WriteXLSXToWriter(&buf, "StreamSheet", nil, rowChan); err != nil {
		panic(err)
	}
	_ = os.WriteFile("stream.xlsx", buf.Bytes(), 0644)
	fmt.Println("XLSX stream written:", "stream.xlsx")
}
