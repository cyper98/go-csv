package main

import (
	"bytes"
	"fmt"
	"os"
	"time"

	go_csv "github.com/aminofox/go-csv"
)

func main() {
	os.MkdirAll("out", 0755)

	fmt.Println("=== go-csv Full Examples ===")
	fmt.Println()

	csvExamples()
	xlsxSingleSheetExamples()
	xlsxMultiSheetExamples()
	styleExamples()
	formulaExamples()
	advancedExamples()
	encryptionExamples()
	zipExamples()

	fmt.Println("=== All Examples Completed ===")
}

func csvExamples() {
	fmt.Println("=== CSV Examples ===")

	headers := []any{"id", "name", "email", "age", "score", "active"}
	data := [][]any{
		{1, "alice", "alice@example.com", 25, 9.5, true},
		{2, "bob", "bob@example.com", 30, 8.1, false},
		{3, "charlie", "charlie@example.com", 28, 7.5, true},
	}

	go_csv.WriteCSVFile("out/users.csv", headers, data, &go_csv.CSVOptions{BOM: true}, nil)
	fmt.Println("written: out/users.csv")

	go_csv.WriteCSVFile("out/users_pipe.csv", []any{"id", "name", "email"}, data, &go_csv.CSVOptions{Comma: '|'}, nil)
	fmt.Println("written: out/users_pipe.csv")

	file, _ := os.Create("out/users_stream.csv")
	defer file.Close()
	rows := make(chan []any, 1024)
	go func() {
		for _, r := range data {
			rows <- r
		}
		close(rows)
	}()
	go_csv.WriteCSVToWriter(file, headers, rows, &go_csv.CSVOptions{BOM: true}, nil)
	fmt.Println("written: out/users_stream.csv")

	h, r, err := go_csv.ReadCSVFile("out/users.csv", &go_csv.CSVReadOptions{HasHeader: true})
	if err != nil {
		fmt.Printf("Read error: %v\n", err)
	} else {
		fmt.Printf("Headers: %v\n", h)
		count := 0
		for row := range r {
			fmt.Printf("Row: %v\n", row)
			count++
			if count >= 3 {
				break
			}
		}
	}

	customToString := func(v any) (string, bool) {
		switch t := v.(type) {
		case float64:
			return fmt.Sprintf("%.2f", t), true
		case time.Time:
			return t.Format("2006-01-02"), true
		case bool:
			if t {
				return "Yes", true
			}
			return "No", true
		default:
			return "", false
		}
	}
	go_csv.WriteCSVFile("out/users_custom.csv", headers, data, nil, customToString)
	fmt.Println("written: out/users_custom.csv")
}

func xlsxSingleSheetExamples() {
	fmt.Println("=== XLSX Single Sheet ===")

	headers := []any{"ID", "Name", "Email", "Age", "Score", "Active"}
	data := [][]any{
		{1, "alice", "alice@example.com", 25, 9.5, true},
		{2, "bob", "bob@example.com", 30, 8.1, false},
		{3, "charlie", "charlie@example.com", 28, 7.5, true},
	}

	go_csv.WriteXLSXFile("out/users.xlsx", "Users", headers, data)
	fmt.Println("written: out/users.xlsx")

	var buf bytes.Buffer
	rowCh := make(chan []any, 1024)
	go func() {
		rowCh <- []any{"A", "B", "C"}
		rowCh <- []any{1, 2, 3}
		rowCh <- []any{4, 5, 6}
		close(rowCh)
	}()
	go_csv.WriteXLSXToWriter(&buf, "StreamSheet", nil, rowCh)
	os.WriteFile("out/stream.xlsx", buf.Bytes(), 0644)
	fmt.Println("written: out/stream.xlsx")
}

func xlsxMultiSheetExamples() {
	fmt.Println("=== XLSX Multi-Sheet ===")

	xw := go_csv.NewXLSXWriter()
	xw.AddSheet("Sales 2023")
	xw.AddSheet("Sales 2024")
	xw.AddSheet("Summary")

	xw.WriteRow("Sales 2023", []any{"Product", "Q1", "Q2", "Q3", "Q4", "Total"})
	xw.WriteRow("Sales 2023", []any{"Widget", 100, 150, 120, 180, 550})
	xw.WriteRow("Sales 2023", []any{"Gadget", 200, 250, 220, 280, 950})

	xw.WriteRow("Sales 2024", []any{"Product", "Q1", "Q2", "Q3", "Q4", "Total"})
	xw.WriteRow("Sales 2024", []any{"Widget", 120, 180, 140, 200, 640})
	xw.WriteRow("Sales 2024", []any{"Gadget", 220, 280, 240, 320, 1060})

	xw.WriteRow("Summary", []any{"Product", "2023 Total", "2024 Total"})
	xw.WriteRow("Summary", []any{"Widget", 550, 640})
	xw.WriteRow("Summary", []any{"Gadget", 950, 1060})

	xw.Close()
	xw.WriteToFile("out/multi_sheet.xlsx")
	fmt.Println("written: out/multi_sheet.xlsx")
}

func styleExamples() {
	fmt.Println("=== Style Examples ===")

	xw := go_csv.NewXLSXWriter()
	xw.AddSheet("Styles")

	xw.WriteRow("Styles", []any{"Product", "Q1", "Q2", "Q3", "Q4", "Total"})
	xw.WriteRow("Styles", []any{"Widget", 100, 150, 120, 180, 550})
	xw.WriteRow("Styles", []any{"Gadget", 200, 250, 220, 280, 950})

	xw.MergeCell("Styles", "A1", "F1")

	xw.Close()
	xw.WriteToFile("out/styles.xlsx")
	fmt.Println("written: out/styles.xlsx")
}

func formulaExamples() {
	fmt.Println("=== Formula Examples ===")

	fc := go_csv.NewFormulaContext()
	fc.SetCellValue("Sheet1", "A1", 100)
	fc.SetCellValue("Sheet1", "A2", 200)
	fc.SetCellValue("Sheet1", "A3", 300)
	fc.SetCellValue("Sheet1", "B1", "Hello")
	fc.SetCellValue("Sheet1", "B2", "World")

	tests := []string{
		"=SUM(A1:A3)",
		"=AVERAGE(A1:A3)",
		"=MAX(A1:A3)",
		"=MIN(A1:A3)",
		"=COUNT(A1:A3)",
		"=IF(A1>50,\"High\",\"Low\")",
		"=AND(A1>50,A2>100)",
		"=OR(A1>200,A2>200)",
		"=LEN(B1)",
		"=UPPER(B1)",
		"=CONCATENATE(B1,\" \",B2)",
	}

	fmt.Println("Formula tests:")
	for _, formula := range tests {
		result := fc.Evaluate(formula)
		if result.Error != nil {
			fmt.Printf("  %s: Error: %v\n", formula, result.Error)
		} else {
			fmt.Printf("  %s = %v\n", formula, result.Value)
		}
	}

	xw := go_csv.NewXLSXWriter()
	xw.AddSheet("Formulas")
	xw.WriteRow("Formulas", []any{100})
	xw.WriteRow("Formulas", []any{200})
	xw.WriteRow("Formulas", []any{300})
	xw.Close()
	xw.WriteToFile("out/formulas.xlsx")
	fmt.Println("written: out/formulas.xlsx")
}

func advancedExamples() {
	fmt.Println("=== Advanced Examples ===")

	xw := go_csv.NewXLSXWriter()
	xw.AddSheet("Advanced")

	freeze := &go_csv.FreezePane{ColSplit: 2, RowSplit: 1, TopLeftCell: "C2"}
	xw.SetFreezePane("Advanced", freeze)

	xw.WriteRow("Advanced", []any{"Name", "Age", "City", "Score"})
	xw.WriteRow("Advanced", []any{"Alice", 25, "NYC", 90})
	xw.WriteRow("Advanced", []any{"Bob", 30, "LA", 85})
	xw.WriteRow("Advanced", []any{"Charlie", 28, "NYC", 88})
	xw.WriteRow("Advanced", []any{"David", 35, "NYC", 92})

	xw.MergeCell("Advanced", "A10", "C10")
	xw.WriteRow("Advanced", []any{"Merged Cell"})

	xw.SetColWidth("Advanced", "A", 20)
	xw.SetRowHeight("Advanced", 1, 30)

	xw.Close()
	xw.WriteToFile("out/advanced.xlsx")
	fmt.Println("written: out/advanced.xlsx")
}

func encryptionExamples() {
	fmt.Println("=== Encryption Examples ===")

	headers := []any{"ID", "Name", "Value"}
	data := [][]any{
		{1, "Item A", 100},
		{2, "Item B", 200},
		{3, "Item C", 300},
	}

	var buf bytes.Buffer
	go_csv.WriteXLSXToWriter(&buf, "Encrypted", headers, toChan(data))
	plainData := buf.Bytes()

	password := "secret123"
	encrypted, err := go_csv.EncryptXLSX(plainData, password)
	if err != nil {
		fmt.Printf("Encryption error: %v\n", err)
		return
	}

	os.WriteFile("out/encrypted.xlsx", encrypted, 0644)
	fmt.Println("written: out/encrypted.xlsx (encrypted)")

	decrypted, err := go_csv.DecryptXLSX(encrypted, password)
	if err != nil {
		fmt.Printf("Decryption error: %v\n", err)
		return
	}

	os.WriteFile("out/decrypted.xlsx", decrypted, 0644)
	fmt.Println("written: out/decrypted.xlsx (decrypted)")
}

func zipExamples() {
	fmt.Println("=== ZIP Examples ===")

	csvData, _ := os.ReadFile("out/users.csv")
	xlsxData, _ := os.ReadFile("out/users.xlsx")

	files := map[string][]byte{
		"csv/users.csv":   csvData,
		"xlsx/users.xlsx": xlsxData,
		"readme.txt":      []byte("Export"),
	}

	go_csv.ZipBuffers("out/export.zip", files)
	fmt.Println("written: out/export.zip")

	go_csv.ZipBuffersParallel("out/export_parallel.zip", files)
	fmt.Println("written: out/export_parallel.zip")

	paths := []string{"out/users.csv", "out/users.xlsx"}
	go_csv.ZipPaths("out/export_paths.zip", paths, "out")
	fmt.Println("written: out/export_paths.zip")

	go_csv.ZipPathsParallel("out/export_paths_parallel.zip", paths, "out")
	fmt.Println("written: out/export_paths_parallel.zip")
}

func toChan(data [][]any) chan []any {
	ch := make(chan []any, len(data))
	go func() {
		for _, r := range data {
			ch <- r
		}
		close(ch)
	}()
	return ch
}
