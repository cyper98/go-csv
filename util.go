package go_csv

import (
	"bytes"
	"encoding/json"
	"fmt"
	"reflect"
	"time"
)

// Convert any -> string: supports number, bool, time, fmt.Stringer, map/slice => JSON
func DefaultToString(v any) (string, bool) {
	if v == nil {
		return "", true
	}

	switch t := v.(type) {
	case string:
		return t, true
	case []byte:
		return string(t), true
	case bool:
		if t {
			return "true", true
		}
		return "false", true
	case time.Time:
		return t.UTC().Format(time.RFC3339), true
	case fmt.Stringer:
		return t.String(), true
	}

	rv := reflect.ValueOf(v)
	switch rv.Kind() {
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		return fmt.Sprintf("%d", rv.Int()), true
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		return fmt.Sprintf("%d", rv.Uint()), true
	case reflect.Float32, reflect.Float64:
		// Giữ nguyên, để CSV/XLSX đọc không mất chính xác; nếu muốn 2 số thập phân tự format sau
		return fmt.Sprintf("%v", rv.Float()), true
	case reflect.Slice, reflect.Array, reflect.Map, reflect.Struct:
		// fallback JSON
		var b bytes.Buffer
		enc := json.NewEncoder(&b)
		enc.SetEscapeHTML(false)
		if err := enc.Encode(v); err == nil {
			s := b.String()
			if len(s) > 0 && s[len(s)-1] == '\n' {
				s = s[:len(s)-1]
			}
			return s, true
		}
	}
	return fmt.Sprintf("%v", v), true
}
