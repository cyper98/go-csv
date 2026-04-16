# go-csv

Lightweight Excel/CSV library for Go - zero dependencies.

## Features

### CSV
- Read & Write
- Streaming (`<-chan []any`)
- Custom delimiter (Comma, Tab, etc.)
- BOM for UTF-8
- CR/LF line endings
- Comment lines support
- Validation functions

### XLSX
- Read & Write (single/multi-sheet)
- Multi-sheet with `NewXLSXWriter()`
- Data types: String, Number, Date/Time
- Column name caching (1024 first columns)
- Rich text support

### Styles
- Font (Bold, Italic, Underline, Size, Color, Name)
- Fill (Pattern, FGColor, BGColor)
- Border (Left, Right, Top, Bottom, Style, Color)
- Alignment (Horizontal, Vertical, Wrap, Rotation, Indent)
- Number formats

### Charts (18+ types)
- Line, Line3D, Column, Column3D
- Bar, Bar3D, Area, Area3D
- Pie, Pie3D, PieOfPie, Doughnut
- Scatter, Radar, RadarArea, RadarMarker
- Surface, Stock
- Grouping (Standard, Stacked, PercentStacked, Clustered)
- Legend, Axis titles, Data labels, Markers

### Pivot Tables
- Row/Column/Data fields
- Grand totals
- Compact/Outline view
- Multiple data fields

### VBA
- Modules, Procedures, Arguments

### Advanced
- Conditional Formatting
- AutoFilter
- Freeze Panes
- Merge Cells
- Data Validation (dropdown, rules)
- Comments
- Tables
- Images (placeholder)
- Encryption/Decryption
- Sheet Protection

### Formulas (30+ functions)
- Math: SUM, AVERAGE, COUNT, COUNTA, MAX, MIN, ABS, ROUND, FLOOR, CEILING, SQRT, POWER, MOD
- Logic: IF, AND, OR, NOT
- Text: LEN, UPPER, LOWER, TRIM, LEFT, RIGHT, MID, CONCATENATE, REPLACE, SUBSTITUTE
- Date: DATE, YEAR, MONTH, DAY, WEEKDAY, TODAY, NOW
- Info: ISBLANK, ISNUMBER, ISTEXT, ISERROR, TRUE, FALSE, N

### ZIP
- Compress buffers/files
- Parallel compression

## Installation

```bash
go get github.com/aminofox/go-csv
```

## Quick Usage

### CSV
```go
go_csv.WriteCSVFile("out.csv", headers, data, &go_csv.CSVOptions{BOM: true}, nil)
headers, rows, _ := go_csv.ReadCSVFile("data.csv", &go_csv.CSVReadOptions{HasHeader: true})
```

### XLSX
```go
xw := go_csv.NewXLSXWriter()
xw.AddSheet("Sheet1")
xw.AddSheet("Sheet2")
xw.WriteRow("Sheet1", []any{"A", "B"})
xw.Close()
xw.WriteToFile("out.xlsx")
```

### Styles
```go
style := &go_csv.Style{
    Font:      &go_csv.Font{Bold: true, Size: 12, Color: "FF0000"},
    Fill:      &go_csv.Fill{Pattern: "solid", FGColor: "FFFF00"},
    Alignment: &go_csv.Alignment{Horizontal: "center"},
}
xw.SetCellStyle("Sheet1", "A1", style)
```

### Charts
```go
chart := &go_csv.Chart{
    Type: go_csv.ChartColumn,
    Title: "Sales",
    Series: []go_csv.ChartSeries{
        {Name: "Q1", Categories: "A1:A4", Values: "B1:B4"},
        {Name: "Q2", Categories: "A1:A4", Values: "C1:C4"},
    },
    Legend:    true,
    LegendPos: "r",
    XAxisLabel: "Month",
    YAxisLabel: "Sales",
}
xw.AddChartFull("Sheet1", "E1", chart)
```

### Formulas
```go
fc := go_csv.NewFormulaContext()
fc.SetCellValue("Sheet1", "A1", 100)
fc.SetCellValue("Sheet1", "A2", 200)
result := fc.Evaluate("=SUM(A1:A2)")
// result.Value = 300
```

### Encryption
```go
encrypted, _ := go_csv.EncryptXLSX(data, "password")
decrypted, _ := go_csv.DecryptXLSX(encrypted, "password")
```

### ZIP
```go
go_csv.ZipBuffers("out.zip", files)
go_csv.ZipBuffersParallel("out.zip", files)
```

## API Reference

### CSV
| Function | Description |
|----------|-------------|
| `WriteCSVFile` | Write CSV file |
| `WriteCSVToWriter` | Write with streaming |
| `ReadCSVFile` | Read CSV file |
| `ReadCSVFromReader` | Read with reader |
| `ReadAllCSV` | Read all to memory |
| `ReadCSVAsAny` | Read as []any |

### XLSX
| Function | Description |
|----------|-------------|
| `WriteXLSXFile` | Write (single sheet) |
| `WriteXLSXToWriter` | Write (streaming) |
| `NewXLSXWriter` | Multi-sheet writer |
| `AddSheet` | Add sheet |
| `WriteRow` | Write row |
| `WriteToFile` | Save to file |
| `SetCellStyle` | Apply style |
| `MergeCell` | Merge cells |
| `AddChartFull` | Add chart |
| `AddPivotTableFull` | Add pivot table |
| `AddCommentFull` | Add comment |
| `SetDataValidationFull` | Add validation |
| `AutoFilterFull` | Add autofilter |
| `SetFreezePane` | Freeze panes |
| `ProtectSheet` | Protect sheet |
| `ReadXLSXFile` | Read XLSX |
| `AddImage` | Add image |

### Formulas
| Function | Description |
|----------|-------------|
| `NewFormulaContext` | Create context |
| `Evaluate` | Evaluate formula |
| `SetCellValue` | Set cell |
| `SetSheetData` | Set sheet data |
| `GetCellValue` | Get cell value |

### ZIP
| Function | Description |
|----------|-------------|
| `ZipBuffers` | Compress buffers |
| `ZipPaths` | Compress files |
| `ZipBuffersParallel` | Parallel compress |
| `ZipPathsParallel` | Parallel compress |

### Encryption
| Function | Description |
|----------|-------------|
| `EncryptXLSX` | Encrypt XLSX |
| `DecryptXLSX` | Decrypt XLSX |

## License

MIT