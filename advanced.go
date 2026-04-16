package go_csv

import (
	"crypto/aes"
	"crypto/cipher"
	"crypto/sha1"
	"fmt"
	"math"
	"math/rand"
	"regexp"
	"strconv"
	"strings"
	"time"
)

type FormulaType int

const (
	FormulaValue FormulaType = iota
	FormulaString
	FormulaBool
)

type FormulaResult struct {
	Type  FormulaType
	Value any
	Error error
}

var currentFormulaContext *FormulaContext

type FuncHandler func(args []any) *FormulaResult

var formulaFunctions = map[string]FuncHandler{
	"SUM": func(args []any) *FormulaResult {
		sum := 0.0
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil {
						if n, ok := toFloat64(val); ok {
							sum += n
						}
					}
				}
			} else if n, ok := toFloat64(a); ok {
				sum += n
			}
		}
		return &FormulaResult{Type: FormulaValue, Value: sum}
	},
	"AVERAGE": func(args []any) *FormulaResult {
		sum, count := 0.0, 0
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil {
						if n, ok := toFloat64(val); ok {
							sum += n
							count++
						}
					}
				}
			} else if n, ok := toFloat64(a); ok {
				sum += n
				count++
			}
		}
		if count == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		return &FormulaResult{Type: FormulaValue, Value: sum / float64(count)}
	},
	"COUNT": func(args []any) *FormulaResult {
		count := 0
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil {
						if _, ok := toFloat64(val); ok {
							count++
						}
					}
				}
			} else if _, ok := toFloat64(a); ok {
				count++
			}
		}
		return &FormulaResult{Type: FormulaValue, Value: float64(count)}
	},
	"COUNTA": func(args []any) *FormulaResult {
		count := 0
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil && val != "" {
						count++
					}
				}
			} else if a != nil && a != "" {
				count++
			}
		}
		return &FormulaResult{Type: FormulaValue, Value: float64(count)}
	},
	"MAX": func(args []any) *FormulaResult {
		max := math.Inf(-1)
		hasValue := false
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil {
						if n, ok := toFloat64(val); ok {
							if n > max {
								max = n
							}
							hasValue = true
						}
					}
				}
			} else if n, ok := toFloat64(a); ok {
				if n > max {
					max = n
				}
				hasValue = true
			}
		}
		if !hasValue {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		return &FormulaResult{Type: FormulaValue, Value: max}
	},
	"MIN": func(args []any) *FormulaResult {
		min := math.Inf(1)
		hasValue := false
		for _, a := range args {
			if cells, ok := a.([]string); ok {
				for _, cell := range cells {
					if val := currentFormulaContext.GetCellValue("Sheet1", cell); val != nil {
						if n, ok := toFloat64(val); ok {
							if n < min {
								min = n
							}
							hasValue = true
						}
					}
				}
			} else if n, ok := toFloat64(a); ok {
				if n < min {
					min = n
				}
				hasValue = true
			}
		}
		if !hasValue {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		return &FormulaResult{Type: FormulaValue, Value: min}
	},
	"IF": func(args []any) *FormulaResult {
		if len(args) < 3 {
			return &FormulaResult{Type: FormulaValue, Value: nil, Error: fmt.Errorf("IF requires 3 arguments")}
		}
		cond, _ := toBool(args[0])
		if cond {
			return &FormulaResult{Type: FormulaValue, Value: args[1]}
		}
		return &FormulaResult{Type: FormulaValue, Value: args[2]}
	},
	"AND": func(args []any) *FormulaResult {
		for _, a := range args {
			if cond, ok := toBool(a); ok && !cond {
				return &FormulaResult{Type: FormulaBool, Value: false}
			}
		}
		return &FormulaResult{Type: FormulaBool, Value: true}
	},
	"OR": func(args []any) *FormulaResult {
		for _, a := range args {
			if cond, ok := toBool(a); ok && cond {
				return &FormulaResult{Type: FormulaBool, Value: true}
			}
		}
		return &FormulaResult{Type: FormulaBool, Value: false}
	},
	"NOT": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaBool, Value: true}
		}
		cond, _ := toBool(args[0])
		return &FormulaResult{Type: FormulaBool, Value: !cond}
	},
	"ABS": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		if n, ok := toFloat64(args[0]); ok {
			return &FormulaResult{Type: FormulaValue, Value: math.Abs(n)}
		}
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	},
	"ROUND": func(args []any) *FormulaResult {
		if len(args) < 2 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		n, _ := toFloat64(args[0])
		digits, _ := toInt(args[1])
		mult := math.Pow(10, float64(digits))
		if n >= 0 {
			return &FormulaResult{Type: FormulaValue, Value: math.Floor(n*mult+0.5) / mult}
		}
		return &FormulaResult{Type: FormulaValue, Value: math.Ceil(n*mult-0.5) / mult}
	},
	"FLOOR": func(args []any) *FormulaResult {
		if len(args) < 2 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		n, _ := toFloat64(args[0])
		sig, _ := toFloat64(args[1])
		if sig == 0 {
			sig = 1
		}
		return &FormulaResult{Type: FormulaValue, Value: math.Floor(n/sig) * sig}
	},
	"CEILING": func(args []any) *FormulaResult {
		if len(args) < 2 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		n, _ := toFloat64(args[0])
		sig, _ := toFloat64(args[1])
		if sig == 0 {
			sig = 1
		}
		return &FormulaResult{Type: FormulaValue, Value: math.Ceil(n/sig) * sig}
	},
	"SQRT": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		if n, ok := toFloat64(args[0]); ok && n >= 0 {
			return &FormulaResult{Type: FormulaValue, Value: math.Sqrt(n)}
		}
		return &FormulaResult{Type: FormulaValue, Value: math.NaN()}
	},
	"POWER": func(args []any) *FormulaResult {
		if len(args) < 2 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		base, _ := toFloat64(args[0])
		exp, _ := toFloat64(args[1])
		return &FormulaResult{Type: FormulaValue, Value: math.Pow(base, exp)}
	},
	"MOD": func(args []any) *FormulaResult {
		if len(args) < 2 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		n, _ := toFloat64(args[0])
		d, _ := toFloat64(args[1])
		if d == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		return &FormulaResult{Type: FormulaValue, Value: math.Mod(n, d)}
	},
	"LEN": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		s := fmt.Sprintf("%v", args[0])
		return &FormulaResult{Type: FormulaValue, Value: float64(len(s))}
	},
	"UPPER": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		if s, ok := args[0].(string); ok {
			return &FormulaResult{Type: FormulaString, Value: strings.ToUpper(s)}
		}
		return &FormulaResult{Type: FormulaString, Value: strings.ToUpper(fmt.Sprintf("%v", args[0]))}
	},
	"LOWER": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		if s, ok := args[0].(string); ok {
			return &FormulaResult{Type: FormulaString, Value: strings.ToLower(s)}
		}
		return &FormulaResult{Type: FormulaString, Value: strings.ToLower(fmt.Sprintf("%v", args[0]))}
	},
	"TRIM": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		return &FormulaResult{Type: FormulaString, Value: strings.TrimSpace(fmt.Sprintf("%v", args[0]))}
	},
	"MID": func(args []any) *FormulaResult {
		if len(args) < 3 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		s := ""
		if str, ok := args[0].(string); ok {
			s = str
		} else {
			s = fmt.Sprintf("%v", args[0])
		}
		start, _ := toInt(args[1])
		length, _ := toInt(args[2])
		start--
		if start < 0 {
			start = 0
		}
		if start >= len(s) {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		end := start + length
		if end > len(s) {
			end = len(s)
		}
		return &FormulaResult{Type: FormulaString, Value: s[start:end]}
	},
	"CONCATENATE": func(args []any) *FormulaResult {
		var parts []string
		for _, a := range args {
			parts = append(parts, fmt.Sprintf("%v", a))
		}
		return &FormulaResult{Type: FormulaString, Value: strings.Join(parts, " ")}
	},
	"REPLACE": func(args []any) *FormulaResult {
		if len(args) < 4 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		s := fmt.Sprintf("%v", args[0])
		start, _ := toInt(args[1])
		length, _ := toInt(args[2])
		replacement := fmt.Sprintf("%v", args[3])
		start--
		if start < 0 {
			start = 0
		}
		if start >= len(s) {
			return &FormulaResult{Type: FormulaString, Value: s}
		}
		end := start + length
		if end > len(s) {
			end = len(s)
		}
		return &FormulaResult{Type: FormulaString, Value: s[:start] + replacement + s[end:]}
	},
	"SUBSTITUTE": func(args []any) *FormulaResult {
		if len(args) < 3 {
			return &FormulaResult{Type: FormulaString, Value: ""}
		}
		s := fmt.Sprintf("%v", args[0])
		old := fmt.Sprintf("%v", args[1])
		new := fmt.Sprintf("%v", args[2])
		_ = true
		if len(args) >= 4 {
			if n, ok := toInt(args[3]); ok && n > 0 {
				s = replaceNth(s, old, new, n)
				return &FormulaResult{Type: FormulaString, Value: s}
			}
		}
		s = strings.ReplaceAll(s, old, new)
		return &FormulaResult{Type: FormulaString, Value: s}
	},
	"DATE": func(args []any) *FormulaResult {
		if len(args) < 3 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		year, _ := toInt(args[0])
		month, _ := toInt(args[1])
		day, _ := toInt(args[2])
		t := time.Date(year, time.Month(month), day, 0, 0, 0, 0, time.UTC)
		excelDate := t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24
		return &FormulaResult{Type: FormulaValue, Value: excelDate}
	},
	"YEAR": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		t := excelTimeToTime(args[0])
		return &FormulaResult{Type: FormulaValue, Value: float64(t.Year())}
	},
	"MONTH": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		t := excelTimeToTime(args[0])
		return &FormulaResult{Type: FormulaValue, Value: float64(t.Month())}
	},
	"DAY": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		t := excelTimeToTime(args[0])
		return &FormulaResult{Type: FormulaValue, Value: float64(t.Day())}
	},
	"WEEKDAY": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 1.0}
		}
		t := excelTimeToTime(args[0])
		return &FormulaResult{Type: FormulaValue, Value: float64(int(t.Weekday()) + 1)}
	},
	"TODAY": func(args []any) *FormulaResult {
		_ = args
		t := time.Now()
		excelDate := t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24
		return &FormulaResult{Type: FormulaValue, Value: excelDate}
	},
	"NOW": func(args []any) *FormulaResult {
		_ = args
		t := time.Now()
		excelDate := t.Sub(time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)).Hours() / 24
		return &FormulaResult{Type: FormulaValue, Value: excelDate}
	},
	"TRUE": func(args []any) *FormulaResult {
		_ = args
		return &FormulaResult{Type: FormulaBool, Value: true}
	},
	"FALSE": func(args []any) *FormulaResult {
		_ = args
		return &FormulaResult{Type: FormulaBool, Value: false}
	},
	"ISBLANK": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaBool, Value: true}
		}
		return &FormulaResult{Type: FormulaBool, Value: args[0] == nil || args[0] == ""}
	},
	"ISNUMBER": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaBool, Value: false}
		}
		_, ok := toFloat64(args[0])
		return &FormulaResult{Type: FormulaBool, Value: ok}
	},
	"ISTEXT": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaBool, Value: false}
		}
		_, ok := args[0].(string)
		return &FormulaResult{Type: FormulaBool, Value: ok}
	},
	"ISERROR": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaBool, Value: false}
		}
		_, ok := args[0].(error)
		if ok {
			return &FormulaResult{Type: FormulaBool, Value: true}
		}
		if n, ok := toFloat64(args[0]); ok {
			return &FormulaResult{Type: FormulaBool, Value: math.IsNaN(n) || math.IsInf(n, 0)}
		}
		return &FormulaResult{Type: FormulaBool, Value: false}
	},
	"N": func(args []any) *FormulaResult {
		if len(args) == 0 {
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		if n, ok := toFloat64(args[0]); ok {
			return &FormulaResult{Type: FormulaValue, Value: n}
		}
		if b, ok := toBool(args[0]); ok {
			if b {
				return &FormulaResult{Type: FormulaValue, Value: 1.0}
			}
			return &FormulaResult{Type: FormulaValue, Value: 0.0}
		}
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	},
}

type FormulaContext struct {
	Data   map[string]map[string]any
	Values map[string]any
}

func NewFormulaContext() *FormulaContext {
	return &FormulaContext{
		Data:   make(map[string]map[string]any),
		Values: make(map[string]any),
	}
}

func (fc *FormulaContext) SetSheetData(sheet string, data map[string]any) {
	fc.Data[sheet] = data
}

func (fc *FormulaContext) GetCellValue(sheet, cell string) any {
	if data, ok := fc.Data[sheet]; ok {
		return data[cell]
	}
	return nil
}

func (fc *FormulaContext) SetCellValue(sheet, cell string, value any) {
	if fc.Data[sheet] == nil {
		fc.Data[sheet] = make(map[string]any)
	}
	fc.Data[sheet][cell] = value
}

func (fc *FormulaContext) Evaluate(formula string) *FormulaResult {
	formula = strings.TrimSpace(formula)
	if !strings.HasPrefix(formula, "=") {
		return &FormulaResult{Type: FormulaValue, Value: formula}
	}
	formula = formula[1:]

	if strings.HasPrefix(formula, "SUM(") {
		return fc.evaluateSum(formula)
	}
	if strings.HasPrefix(formula, "AVERAGE(") {
		return fc.evaluateAverage(formula)
	}
	if strings.HasPrefix(formula, "IF(") {
		return fc.evaluateIf(formula)
	}

	for name, handler := range formulaFunctions {
		prefix := name + "("
		if strings.HasPrefix(strings.ToUpper(formula), prefix) {
			currentFormulaContext = fc
			args := fc.parseArguments(strings.TrimSuffix(strings.TrimPrefix(formula, prefix), ")"))
			defer func() { currentFormulaContext = nil }()
			return handler(args)
		}
	}

	if strings.Contains(formula, "+") {
		return fc.evaluateBinaryOp(formula, "+", func(a, b float64) float64 { return a + b })
	}
	if strings.Contains(formula, "-") {
		return fc.evaluateBinaryOp(formula, "-", func(a, b float64) float64 { return a - b })
	}
	if strings.Contains(formula, "*") {
		return fc.evaluateBinaryOp(formula, "*", func(a, b float64) float64 { return a * b })
	}
	if strings.Contains(formula, "/") {
		result := fc.evaluateBinaryOp(formula, "/", func(a, b float64) float64 {
			if b == 0 {
				return math.NaN()
			}
			return a / b
		})
		return result
	}
	if strings.Contains(formula, "^") {
		return fc.evaluateBinaryOp(formula, "^", math.Pow)
	}

	if strings.Contains(formula, "&") {
		return fc.evaluateConcat(formula)
	}

	if strings.Contains(formula, "=") {
		return fc.evaluateComparison(formula, "=")
	}
	if strings.Contains(formula, "<>") {
		return fc.evaluateComparison(formula, "<>")
	}
	if strings.Contains(formula, "<=") {
		return fc.evaluateComparison(formula, "<=")
	}
	if strings.Contains(formula, ">=") {
		return fc.evaluateComparison(formula, ">=")
	}
	if strings.Contains(formula, "<") {
		return fc.evaluateComparison(formula, "<")
	}
	if strings.Contains(formula, ">") {
		return fc.evaluateComparison(formula, ">")
	}

	return &FormulaResult{Type: FormulaString, Value: formula}
}

func (fc *FormulaContext) parseArguments(argsStr string) []any {
	if argsStr == "" {
		return []any{}
	}

	var args []any
	var current strings.Builder
	depth := 0
	inString := false

	for i := 0; i < len(argsStr); i++ {
		c := rune(argsStr[i])

		if c == '"' && (i == 0 || argsStr[i-1] != '\\') {
			inString = !inString
			current.WriteRune(c)
			continue
		}

		if inString {
			current.WriteRune(c)
			continue
		}

		switch c {
		case '(':
			depth++
			current.WriteRune(c)
		case ')':
			depth--
			current.WriteRune(c)
		case ',':
			if depth == 0 {
				arg := strings.TrimSpace(current.String())
				if arg != "" {
					if fc.isCellRange(arg) {
						args = append(args, fc.getCellsInRangeFromStr(arg))
					} else if strings.HasPrefix(arg, "\"") && strings.HasSuffix(arg, "\"") {
						args = append(args, strings.Trim(arg, "\""))
					} else {
						args = append(args, fc.evaluateSingleValue(arg))
					}
				}
				current.Reset()
			} else {
				current.WriteRune(c)
			}
		default:
			current.WriteRune(c)
		}
	}

	arg := strings.TrimSpace(current.String())
	if arg != "" {
		if fc.isCellRange(arg) {
			args = append(args, fc.getCellsInRangeFromStr(arg))
		} else if strings.HasPrefix(arg, "\"") && strings.HasSuffix(arg, "\"") {
			args = append(args, strings.Trim(arg, "\""))
		} else {
			args = append(args, fc.evaluateSingleValue(arg))
		}
	}

	return args
}

func (fc *FormulaContext) isCellRange(s string) bool {
	s = strings.TrimSpace(s)
	re := regexp.MustCompile(`^[A-Z]+\d+:[A-Z]+\d+$`)
	return re.MatchString(strings.ToUpper(s))
}

func (fc *FormulaContext) getCellsInRangeFromStr(rangeStr string) []string {
	parts := strings.Split(rangeStr, ":")
	if len(parts) != 2 {
		return nil
	}
	start := strings.TrimSpace(parts[0])
	end := strings.TrimSpace(parts[1])
	return fc.getCellsInRange(start, end)
}

func (fc *FormulaContext) evaluateSingleValue(arg string) any {
	arg = strings.TrimSpace(arg)

	if strings.HasPrefix(arg, "\"") && strings.HasSuffix(arg, "\"") {
		return strings.Trim(arg, "\"")
	}

	hasComparison := strings.Contains(arg, ">") || strings.Contains(arg, "<") || strings.Contains(arg, "=") || strings.Contains(arg, "<>") || strings.Contains(arg, "<=") || strings.Contains(arg, ">=")
	if hasComparison {
		result := fc.evaluateComparisonExpr(arg)
		return result.Value
	}

	if fc.isCellRange(arg) {
		return fc.getCellsInRangeFromStr(arg)
	}

	if n, ok := toFloat64(arg); ok {
		return n
	}

	cellRefResult := parseCellRef(arg, fc)
	if cellRefResult != nil && arg == strings.TrimSpace(arg) {
		return cellRefResult
	}

	if b, ok := toBool(arg); ok {
		return b
	}

	return cellRefResult
}

func (fc *FormulaContext) evaluateComparisonExpr(expr string) *FormulaResult {
	ops := []string{"<>", "<=", ">=", "=", "<", ">"}

	for _, op := range ops {
		if strings.Contains(expr, op) {
			parts := strings.SplitN(expr, op, 2)
			if len(parts) == 2 {
				left := strings.TrimSpace(parts[0])
				right := strings.TrimSpace(parts[1])

				a := parseCellRef(left, fc)
				b := parseCellRef(right, fc)

				an := toFloat64Explicit(a)
				bn := toFloat64Explicit(b)

				if an != nil && bn != nil {
					switch op {
					case "=":
						return &FormulaResult{Type: FormulaBool, Value: *an == *bn}
					case "<>":
						return &FormulaResult{Type: FormulaBool, Value: *an != *bn}
					case "<":
						return &FormulaResult{Type: FormulaBool, Value: *an < *bn}
					case ">":
						return &FormulaResult{Type: FormulaBool, Value: *an > *bn}
					case "<=":
						return &FormulaResult{Type: FormulaBool, Value: *an <= *bn}
					case ">=":
						return &FormulaResult{Type: FormulaBool, Value: *an >= *bn}
					}
				}

				as := fmt.Sprintf("%v", a)
				bs := fmt.Sprintf("%v", b)
				switch op {
				case "=":
					return &FormulaResult{Type: FormulaBool, Value: as == bs}
				case "<>":
					return &FormulaResult{Type: FormulaBool, Value: as != bs}
				}
			}
		}
	}

	return &FormulaResult{Type: FormulaBool, Value: false}
}

func parseCellRef(ref string, fc *FormulaContext) any {
	ref = strings.ToUpper(ref)
	re := regexp.MustCompile(`^([A-Z]+)(\d+)(:([A-Z]+)(\d+))?$`)
	matches := re.FindStringSubmatch(ref)

	if matches == nil {
		return ref
	}

	if matches[3] == "" {
		useFc := fc
		if useFc == nil {
			useFc = currentFormulaContext
		}
		if useFc != nil {
			return useFc.GetCellValue("Sheet1", ref)
		}
		return nil
	}

	return nil
}

func (fc *FormulaContext) evaluateSum(formula string) *FormulaResult {
	start, end := parseRangeInFormula(formula)
	if start == "" || end == "" {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	sum := 0.0
	for _, cell := range fc.getCellsInRange(start, end) {
		if val := fc.GetCellValue("Sheet1", cell); val != nil {
			if n, ok := toFloat64(val); ok {
				sum += n
			}
		}
	}

	return &FormulaResult{Type: FormulaValue, Value: sum}
}

func (fc *FormulaContext) evaluateAverage(formula string) *FormulaResult {
	start, end := parseRangeInFormula(formula)
	if start == "" || end == "" {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	sum, count := 0.0, 0
	for _, cell := range fc.getCellsInRange(start, end) {
		if val := fc.GetCellValue("Sheet1", cell); val != nil {
			if n, ok := toFloat64(val); ok {
				sum += n
				count++
			}
		}
	}

	if count == 0.0 {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	return &FormulaResult{Type: FormulaValue, Value: sum / float64(count)}
}

func (fc *FormulaContext) evaluateIf(formula string) *FormulaResult {
	args := fc.parseArguments(strings.TrimSuffix(strings.TrimPrefix(formula, "IF("), ")"))
	if len(args) < 3 {
		return &FormulaResult{Type: FormulaValue, Value: nil}
	}

	cond, _ := toBool(args[0])
	if cond {
		return &FormulaResult{Type: FormulaValue, Value: args[1]}
	}
	return &FormulaResult{Type: FormulaValue, Value: args[2]}
}

func (fc *FormulaContext) evaluateBinaryOp(formula string, op string, fn func(float64, float64) float64) *FormulaResult {
	parts := strings.Split(formula, op)
	if len(parts) != 2 {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	a := parseCellRefOrNumber(strings.TrimSpace(parts[0]), fc)
	b := parseCellRefOrNumber(strings.TrimSpace(parts[1]), fc)

	if a == nil || b == nil {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	av := toFloat64Explicit(a)
	bv := toFloat64Explicit(b)

	if av == nil || bv == nil {
		return &FormulaResult{Type: FormulaValue, Value: 0.0}
	}

	return &FormulaResult{Type: FormulaValue, Value: fn(*av, *bv)}
}

func (fc *FormulaContext) evaluateConcat(formula string) *FormulaResult {
	parts := strings.Split(formula, "&")
	var result strings.Builder

	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		if val := parseCellRefOrString(part, fc); val != nil {
			result.WriteString(fmt.Sprintf("%v", val))
		}
	}

	return &FormulaResult{Type: FormulaString, Value: result.String()}
}

func (fc *FormulaContext) evaluateComparison(formula string, op string) *FormulaResult {
	parts := strings.Split(formula, op)
	if len(parts) != 2 {
		return &FormulaResult{Type: FormulaBool, Value: false}
	}

	a := parseCellRefOrString(strings.TrimSpace(parts[0]), fc)
	b := parseCellRefOrString(strings.TrimSpace(parts[1]), fc)

	if a == nil || b == nil {
		return &FormulaResult{Type: FormulaBool, Value: false}
	}

	an := toFloat64Explicit(a)
	bn := toFloat64Explicit(b)

	if an != nil && bn != nil {
		switch op {
		case "=":
			return &FormulaResult{Type: FormulaBool, Value: *an == *bn}
		case "<>":
			return &FormulaResult{Type: FormulaBool, Value: *an != *bn}
		case "<":
			return &FormulaResult{Type: FormulaBool, Value: *an < *bn}
		case ">":
			return &FormulaResult{Type: FormulaBool, Value: *an > *bn}
		case "<=":
			return &FormulaResult{Type: FormulaBool, Value: *an <= *bn}
		case ">=":
			return &FormulaResult{Type: FormulaBool, Value: *an >= *bn}
		}
	}

	as := fmt.Sprintf("%v", a)
	bs := fmt.Sprintf("%v", b)

	switch op {
	case "=":
		return &FormulaResult{Type: FormulaBool, Value: as == bs}
	case "<>":
		return &FormulaResult{Type: FormulaBool, Value: as != bs}
	}

	return &FormulaResult{Type: FormulaBool, Value: false}
}

func parseRangeInFormula(formula string) (string, string) {
	formula = strings.TrimSuffix(formula, ")")
	formula = strings.TrimPrefix(formula, "SUM(")
	formula = strings.TrimPrefix(formula, "AVERAGE(")
	formula = strings.TrimPrefix(formula, "IF(")

	parts := strings.Split(formula, ":")
	if len(parts) != 2 {
		return "", ""
	}

	return strings.TrimSpace(parts[0]), strings.TrimSpace(parts[1])
}

func (fc *FormulaContext) getCellsInRange(start, end string) []string {
	var cells []string

	startCol, startRow := splitCellRef(start)
	endCol, endRow := splitCellRef(end)

	startColNum := columnNameToNumber(startCol)
	endColNum := columnNameToNumber(endCol)

	for row := startRow; row <= endRow; row++ {
		for col := startColNum; col <= endColNum; col++ {
			cells = append(cells, columnNumberToName(col)+strconv.Itoa(row))
		}
	}

	return cells
}

func splitCellRef(ref string) (string, int) {
	re := regexp.MustCompile(`^([A-Z]+)(\d+)$`)
	matches := re.FindStringSubmatch(ref)
	if matches == nil {
		return "", 0
	}
	row, _ := strconv.Atoi(matches[2])
	return matches[1], row
}

func splitCellRefParts(ref string) (string, int, string, int) {
	re := regexp.MustCompile(`^([A-Z]+)(\d+):([A-Z]+)(\d+)$`)
	matches := re.FindStringSubmatch(ref)
	if matches == nil {
		return "", 0, "", 0
	}
	startRow, _ := strconv.Atoi(matches[2])
	endRow, _ := strconv.Atoi(matches[4])
	return matches[1], startRow, matches[3], endRow
}

func parseCellRefOrNumber(s string, fc *FormulaContext) any {
	if n, ok := toFloat64(s); ok {
		return n
	}
	if fc != nil {
		return fc.GetCellValue("Sheet1", s)
	}
	return s
}

func parseCellRefOrString(s string, fc *FormulaContext) any {
	if fc != nil {
		if val := fc.GetCellValue("Sheet1", s); val != nil {
			return val
		}
	}
	return strings.Trim(s, "\"")
}

func toFloat64Explicit(v any) *float64 {
	if n, ok := toFloat64(v); ok {
		return &n
	}
	return nil
}

func toFloat64(v any) (float64, bool) {
	switch t := v.(type) {
	case float64:
		return t, true
	case float32:
		return float64(t), true
	case int:
		return float64(t), true
	case int32:
		return float64(t), true
	case int64:
		return float64(t), true
	case uint:
		return float64(t), true
	case uint32:
		return float64(t), true
	case uint64:
		return float64(t), true
	case string:
		n, err := strconv.ParseFloat(t, 64)
		return n, err == nil
	default:
		n, ok := v.(float64)
		return n, ok
	}
}

func toBool(v any) (bool, bool) {
	switch t := v.(type) {
	case bool:
		return t, true
	case string:
		lower := strings.ToLower(t)
		return lower == "true" || t == "1", true
	case float64:
		return t != 0, true
	case int:
		return t != 0, true
	default:
		if s, ok := v.(string); ok {
			return strings.ToLower(s) == "true" || s == "1", true
		}
		return false, false
	}
}

func toInt(v any) (int, bool) {
	switch t := v.(type) {
	case int:
		return t, true
	case int32:
		return int(t), true
	case int64:
		return int(t), true
	case float64:
		return int(t), true
	case string:
		n, err := strconv.Atoi(t)
		return n, err == nil
	default:
		return 0, false
	}
}

func toFloat32(v any) (float64, bool) {
	switch t := v.(type) {
	case float32:
		return float64(t), true
	case float64:
		return t, true
	case int:
		return float64(t), true
	default:
		if n, ok := v.(float64); ok {
			return n, true
		}
		return 0, false
	}
}

func replaceNth(s, old, new string, n int) string {
	if n <= 0 {
		return s
	}

	result := s
	for i := 0; i < n-1; i++ {
		idx := strings.Index(result, old)
		if idx == -1 {
			break
		}
		result = result[:idx] + result[idx+len(old):]
	}

	return strings.Replace(result, old, new, 1)
}

func excelTimeToTime(v any) time.Time {
	n, ok := toFloat64(v)
	if !ok {
		return time.Time{}
	}

	serial := int64(n)
	nanoseconds := (n - float64(serial)) * 24 * 60 * 60 * 1000000000

	baseDate := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := baseDate.Add(time.Duration(serial) * 24 * time.Hour)
	t = t.Add(time.Duration(nanoseconds))

	return t
}

func (fc *FormulaContext) evaluateCellRef(cell string) any {
	if fc == nil {
		return nil
	}
	return fc.GetCellValue("Sheet1", cell)
}

func columnNameToNumber(name string) int {
	result := 0
	for _, c := range strings.ToUpper(name) {
		result = result*26 + int(c-'A'+1)
	}
	return result
}

func columnNumberToName(num int) string {
	result := ""
	for num > 0 {
		num--
		result = string(rune('A'+num%26)) + result
		num /= 26
	}
	return result
}

type PivotTableOptions struct {
	Name      string
	Range     string
	DestSheet string
	Columns   []string
	Rows      []string
	Values    []string
	Function  string
}

func (pt *PivotTableOptions) ToXML() string {
	xml := fmt.Sprintf(`<pivotTableDefinition name="%s" dest="%s" dataRange="%s"`, pt.Name, pt.DestSheet, pt.Range)
	xml += "><pivotFields>"

	for i := 0; i < 10; i++ {
		xml += "<pivotField/>"
	}

	xml += "</pivotFields></pivotTableDefinition>"
	return xml
}

type ConditionalFormatType string

const (
	CondFormatCellValue  ConditionalFormatType = "cellValue"
	CondFormatExpression ConditionalFormatType = "expression"
	CondFormatColorScale ConditionalFormatType = "colorScale"
	CondFormatDataBar    ConditionalFormatType = "dataBar"
	CondFormatIconSet    ConditionalFormatType = "iconSet"
)

type ConditionalFormatRule struct {
	Type       ConditionalFormatType
	Formula    []string
	Priority   int
	StopIfTrue bool

	CellValue rule   `xml:"rule"`
	Icon      string `xml:"icon"`
	Color     string `xml:"color"`
}

type rule struct {
	Operation string
	Formula   string
}

func (cf *ConditionalFormatRule) ToXML() string {
	var xml strings.Builder
	xml.WriteString(fmt.Sprintf(`<cfRule type="%s" priority="%d" stopIfTrue="%t">`, cf.Type, cf.Priority, cf.StopIfTrue))

	if len(cf.Formula) > 0 {
		xml.WriteString(fmt.Sprintf(`<formula>%s</formula>`, cf.Formula[0]))
	}

	xml.WriteString("</cfRange>")

	return xml.String()
}

type FreezePane struct {
	ColSplit    int
	RowSplit    int
	TopLeftCell string
	Pane        string
}

func (fp *FreezePane) ToXML() string {
	return fmt.Sprintf(`<pane topLeftCell="%s" xSplit="%d" ySplit="%d" activePane="bottomRight" state="frozen"/>`,
		fp.TopLeftCell, fp.ColSplit, fp.RowSplit)
}

func (xw *XLSXWriter) SetFreezePane(sheetName string, freeze *FreezePane) error {
	sw, ok := xw.sheets[sheetName]
	if !ok {
		return ErrSheetNotFound
	}

	_, err := fmt.Fprint(sw.sheetWriter, freeze.ToXML())
	return err
}

func (xw *XLSXWriter) AddConditionalFormat(sheetName string, rangeRef string, cf *ConditionalFormatRule) error {
	return nil
}

func (xw *XLSXWriter) AddPivotTable(sheetName string, pt *PivotTableOptions) error {
	return nil
}

type Macro struct {
	Name    string
	Content []byte
	Type    string
}

func (xw *XLSXWriter) AddVBAProject(macros []byte) error {
	return nil
}

func EncryptXLSX(data []byte, password string) ([]byte, error) {
	if password == "" {
		return nil, fmt.Errorf("password required")
	}

	hash := sha1.Sum([]byte(password))
	key := make([]byte, 16)
	copy(key, hash[:16])

	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}

	iv := make([]byte, aes.BlockSize)
	for i := range iv {
		iv[i] = byte(rand.Int63() & 0xFF)
	}

	stream := cipher.NewCFBEncrypter(block, iv)
	stream.XORKeyStream(data, data)

	result := make([]byte, len(iv)+len(data))
	copy(result, iv)
	copy(result[len(iv):], data)

	return result, nil
}

func DecryptXLSX(data []byte, password string) ([]byte, error) {
	if password == "" {
		return nil, fmt.Errorf("password required")
	}

	if len(data) < aes.BlockSize {
		return nil, fmt.Errorf("invalid encrypted data")
	}

	hash := sha1.Sum([]byte(password))
	key := make([]byte, 16)
	copy(key, hash[:16])

	block, err := aes.NewCipher(key)
	if err != nil {
		return nil, err
	}

	iv := data[:aes.BlockSize]
	encryptedData := data[aes.BlockSize:]

	stream := cipher.NewCFBDecrypter(block, iv)
	stream.XORKeyStream(encryptedData, encryptedData)

	return encryptedData, nil
}

type ImageOptions struct {
	Width           float64
	Height          float64
	OffsetX         float64
	OffsetY         float64
	AltText         string
	LockAspectRatio bool
	PrintObject     bool
}

func (xw *XLSXWriter) AddImage(sheetName, cell string, imageData []byte, opts *ImageOptions) error {
	if len(imageData) >= 4 {
		if imageData[0] == 0xFF && imageData[1] == 0xD8 {
			_ = "jpg"
		}
	}

	_ = sheetName
	_ = cell
	_ = opts

	return nil
}

func (xw *XLSXWriter) GetImage(sheetName, cell string) ([]byte, string, error) {
	return nil, "", nil
}

type TableOptions struct {
	Name       string
	Range      string
	StyleName  string
	ShowHeader bool
	ShowTotals bool
	AutoFilter bool
}

type SheetProtection struct {
	Password  string
	Sheet     bool
	Contents  bool
	Objects   bool
	Scenarios bool
}

func (xw *XLSXWriter) ProtectSheet(sheetName string, protection *SheetProtection) error {
	return nil
}

func (xw *XLSXWriter) UnprotectSheet(sheetName string, password string) error {
	return nil
}

var _ = rand.Int63

func mathRound(x float64, digits int) float64 {
	mult := math.Pow(10, float64(digits))
	if x >= 0 {
		return math.Floor(x*mult+0.5) / mult
	}
	return math.Ceil(x*mult-0.5) / mult
}
