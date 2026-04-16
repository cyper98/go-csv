package go_csv

import (
	"os"
	"testing"
)

func BenchmarkCSVWrite_10kRows(b *testing.B) {
	headers := []any{"id", "name", "email", "age", "score", "active"}
	data := generateTestRows(10000)

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		f, _ := os.CreateTemp("", "bench-*.csv")
		defer os.Remove(f.Name())
		WriteCSVFile(f.Name(), headers, data, &CSVOptions{BOM: true}, nil)
	}
}

func BenchmarkCSVWriteStreaming_10kRows(b *testing.B) {
	headers := []any{"id", "name", "email", "age", "score", "active"}
	data := generateTestRows(10000)

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		f, _ := os.CreateTemp("", "bench-*.csv")
		defer os.Remove(f.Name())

		rows := make(chan []any, 2048)
		go func() {
			for _, r := range data {
				rows <- r
			}
			close(rows)
		}()
		WriteCSVToWriter(f, headers, rows, &CSVOptions{BOM: true}, nil)
	}
}

func BenchmarkXLSXWrite_1kRows(b *testing.B) {
	headers := []any{"id", "name", "email", "age", "score", "active"}
	data := generateTestRows(1000)

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		f, _ := os.CreateTemp("", "bench-*.xlsx")
		defer os.Remove(f.Name())
		WriteXLSXFile(f.Name(), "Sheet1", headers, data)
	}
}

func BenchmarkXLSXMultiSheet_500x2Sheets(b *testing.B) {
	data := generateTestRows(500)

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		f, _ := os.CreateTemp("", "bench-*.xlsx")
		defer os.Remove(f.Name())

		xw := NewXLSXWriter()
		xw.AddSheet("Sheet1")
		xw.AddSheet("Sheet2")

		for _, row := range data {
			xw.WriteRow("Sheet1", row)
			xw.WriteRow("Sheet2", row)
		}

		xw.Close()
		xw.WriteToFile(f.Name())
	}
}

func BenchmarkZipCompress_10Files(b *testing.B) {
	files := make(map[string][]byte, 10)
	content := make([]byte, 1024)
	for i := range content {
		content[i] = byte('a' + i%26)
	}
	for i := 0; i < 10; i++ {
		files[string(rune('a'+rune(i)))+".txt"] = content
	}

	b.ResetTimer()
	b.ReportAllocs()

	for i := 0; i < b.N; i++ {
		f, _ := os.CreateTemp("", "bench-*.zip")
		defer os.Remove(f.Name())
		ZipBuffers(f.Name(), files)
	}
}

func BenchmarkColNameCache(b *testing.B) {
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		for c := 1; c <= 1000; c++ {
			_ = getColName(c)
		}
	}
}

func generateTestRows(count int) [][]any {
	result := make([][]any, count)
	for i := 0; i < count; i++ {
		result[i] = []any{
			i + 1,
			"user-" + string(rune('a'+rune(i%26))),
			"user" + string(rune('a'+rune(i%26))) + "@example.com",
			20 + i%50,
			float64(100+i) / 10.0,
			i%2 == 0,
		}
	}
	return result
}
