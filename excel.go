package excel

import (
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/WedgeNix/xls"
	"github.com/tealeg/xlsx"
)

var layouts = [...]string{
	time.ANSIC,
	time.UnixDate,
	time.RubyDate,
	time.RFC822,
	time.RFC822Z,
	time.RFC850,
	time.RFC1123,
	time.RFC1123Z,
	time.RFC3339,
	time.RFC3339Nano,
	time.Kitchen,
	time.Stamp,
	time.StampMilli,
	time.StampMicro,
	time.StampNano,
	"2006-01-02 15:04:05",
	"2006/1/_2",
	"2006/01/02",
	"1/_2/2006",
	"01/02/2006",
	"2006-1-_2",
	"2006-01-02",
	"1-_2-2006",
	"01-02-2006",
	"01-02-06",
	"1-_2-06",
	"01/02/06",
	"1/_2/06",
}

type Error struct {
	Func string
	Msg  string
	Err  error
}

func (e *Error) Error() string {
	return "excel." + e.Func + ": " + e.Msg + ": " + e.Err.Error()
}

var labelEx = regexp.MustCompile(`[A-Za-z]+`)

var excel = File{
	Comma: ',',
}

// File is the generic Excel API controller.
type File struct {
	initOnce sync.Once

	Name  string
	Comma rune
	// raw sheet daya if no file.
	Sheets [][][]string
	keys   []string
	body   [][]string
}

func (file *File) init() error {
	var initErr error
	file.initOnce.Do(func() {
		sheets := file.Sheets
		if len(sheets) < 0 {
			if len(file.Name) == 0 {
				initErr = errors.New("no file name")
				return
			}
			if file.Comma == 0 {
				file.Comma = excel.Comma
			}

			name, err := regef(file.Name)
			if err != nil {
				initErr = err
				return
			}

			ext := strings.ToLower(filepath.Ext(name))
			switch ext {
			case ".xlsx":
				c, err := xlsx.FileToSlice(name)
				if err != nil {
					initErr = err
					return
				}
				sheets = c
			case ".csv", ".txt":
				f, err := os.Open(name)
				if err != nil {
					initErr = err
					return
				}
				r := csv.NewReader(f)
				r.Comma = file.Comma
				r.LazyQuotes = true
				var records [][]string
				for {
					var record []string
					record, err = r.Read()
					if err == io.EOF {
						break
					}
					if err != nil && !strings.Contains(err.Error(), csv.ErrFieldCount.Error()) {
						initErr = err
						return
					}
					records = append(records, record)
				}
				f.Close()
				sheets = [][][]string{records}
			case ".xls":
				f, err := xls.Open(name, "")
				if err != nil {
					initErr = err
					return
				}
				sheets = [][][]string{f.ReadAllCells()}
			default:
				initErr = errors.New("'" + ext + "' is not an Excel format")
				return
			}
		}

		var (
			bigKeys [][]string
			bigBody [][][]string
		)
		for _, rows := range sheets {
			var (
				longest int
				keys    []string
				body    [][]string
			)
			for i, cells := range rows {
				if len(cells) <= longest {
					continue
				}
				isKeys := true
				for _, cell := range cells {
					if !labelEx.MatchString(cell) {
						isKeys = false
						break
					}
				}
				if !isKeys {
					continue
				}
				longest = len(cells)
				keys = cells
				body = rows[i+1:]
			}
			bigKeys = append(bigKeys, keys)
			bigBody = append(bigBody, body)
		}
		var (
			keys []string
			body [][]string
		)
		if len(bigKeys) > 0 {
			keys = bigKeys[0]
		}
		for _, bod := range bigBody {
			body = append(body, bod...)
		}

		switch {
		case len(keys) == 0:
			initErr = errors.New("empty file")
		case len(body) == 0:
			initErr = errors.New("no values in file")
		default:
			file.keys = keys
			file.body = body
		}
	})
	return initErr
}

// Unmarshal2 parses the excel-encoded data and stores the result
// in the value pointed to by ptr.
// func (file *File) Unmarshal2(ptr interface{}, opt ...interface{}) error {
// 	if err := file.init(); err != nil {
// 		return &Error{"Unmarshal", "init", err}
// 	}

// 	// filter out string options
// 	strs := layouts[:]
// 	for _, o := range opt {
// 		if str, ok := o.(string); ok {
// 			strs = append(strs, str)
// 		}
// 	}

// 	//
// 	//

// 	rptr := reflect.ValueOf(ptr)
// 	if rptr.Kind() != reflect.Ptr {
// 		panic("not a pointer")
// 	}
// 	rv := reflect.Indirect(rptr)
// 	rt := rv.Type()
// 	rmapt := rt
// 	rslt := rt
// 	rstructt := rt

// 	var (
// 		IsMap   bool
// 		IsSlice bool
// 	)

// 	switch rt.Kind() {
// 	case reflect.Map:
// 		IsMap = true

// 		rslt = rmapt.Elem()
// 		switch rslt.Kind() {
// 		case reflect.Slice:
// 			IsSlice = true

// 			rstructt = rslt.Elem()

// 		case reflect.Struct:
// 			rstructt = rslt

// 		default:
// 			panic("not a slice or struct")
// 		}

// 	case reflect.Slice:
// 		IsSlice = true

// 		rmapt = reflect.MapOf(reflect.TypeOf(""), rslt)

// 	case reflect.Struct:
// 		rslt = reflect.SliceOf(rt)
// 		rmapt = reflect.MapOf(reflect.TypeOf(""), rslt)

// 	default:
// 		panic("not a map, slice or struct")
// 	}

// 	//
// 	//

// 	cols := make([]int, rstructt.NumField())
// 	for i := 0; i < len(cols); i++ {
// 		col, err := abbrev(rstructt.Field(i).Name, file.keys...)
// 		if err != nil {
// 			return &Error{"Unmarshal", "abbreviation", err}
// 		}
// 		cols[i] = col
// 	}

// 	if !IsMap && (IsSlice || IsStruct) {
// 		rsl := reflect.New(rslt).Elem()
// 		flds := map[int]byte{}
// 		rfirst := reflect.New(rstructt).Elem()

// 		for _, ln := range file.body {
// 			rstruct := reflect.New(rstructt).Elem()

// 			for f, col := range cols {
// 				rfld := rstruct.Field(f)
// 				rfld2, err := parse(rfld.Type(), ln[col])
// 				if err != nil {
// 					return &Error{"Unmarshal", "parsing data", err}
// 				}
// 				if IsStruct {
// 					if len(flds) == len(cols) {
// 						rv.Set(rfirst)
// 						return nil
// 					}
// 					if _, found := flds[f]; len(ln[col]) == 0 || found {
// 						continue
// 					}
// 					rfirst.Field(f).Set(rfld2)
// 					flds[f] = 0
// 				} else {
// 					rfld.Set(rfld2)
// 				}
// 			}

// 			if IsSlice {
// 				rsl = reflect.Append(rsl, rstruct)
// 			}
// 		}

// 		rv.Set(rsl)
// 		return nil
// 	}

// 	rkeyt := rt.Key()
// 	rmap := reflect.MakeMap(rt)

// 	key := rslt.Name()
// 	if len(key) == 0 {
// 		key = rstructt.Name()
// 	}
// 	keyCol, err := abbrev(key, file.keys...)
// 	if err != nil {
// 		return &Error{"Unmarshal", "abbreviation", err}
// 	}

// 	for _, ln := range file.body {
// 		if keyCol >= len(ln) {
// 			continue
// 		}

// 		rstruct := reflect.New(rstructt).Elem()

// 		key := ln[keyCol]
// 		rkey, err := parse(rkeyt, key)
// 		if err != nil {
// 			return &Error{"Unmarshal", "parsing key", err}
// 		}

// 		for f, col := range cols {
// 			rfld := rstruct.Field(f)
// 			rfld2, err := parse(rfld.Type(), ln[col])
// 			if err != nil {
// 				return &Error{"Unmarshal", "parsing data", err}
// 			}
// 			rfld.Set(rfld2)
// 		}

// 		rmapsl2 := rmap.MapIndex(rkey)
// 		if rmapsl2.Kind() == reflect.Invalid {
// 			rmapsl2 = reflect.New(rslt).Elem()
// 		}
// 		rmap.SetMapIndex(rkey, reflect.Append(rmapsl2, rstruct))
// 	}

// 	// rv.Set(rmap)

// 	return nil
// }

// Unmarshal parses the excel-encoded data and stores the result
// in the value pointed to by ptr.
func (file *File) Unmarshal(ptr interface{}, opt ...interface{}) error {
	if err := file.init(); err != nil {
		return &Error{"Unmarshal", "init", err}
	}

	// filter out string options
	strs := layouts[:]
	for _, o := range opt {
		if str, ok := o.(string); ok {
			strs = append(strs, str)
		}
	}

	//
	//

	rptr := reflect.ValueOf(ptr)
	if rptr.Kind() != reflect.Ptr {
		panic("not a pointer")
	}
	rv := reflect.Indirect(rptr)
	rt := rv.Type()

	var IsMap, IsSlice, IsStruct bool
	rslt := rt
	switch rt.Kind() {
	case reflect.Map:
		IsMap = true
		rslt = rv.Type().Elem()
	case reflect.Slice:
		IsSlice = true
	case reflect.Struct:
		IsStruct = true
	default:
		panic("not a map, slice, or struct")
	}

	if IsStruct {
		rslt = reflect.SliceOf(rt)
	}

	rstructt := rslt
	switch rslt.Kind() {
	case reflect.Slice:
		rstructt = rslt.Elem()
		IsSlice = true
	case reflect.Struct:
		IsStruct = true
	default:
		panic("not a slice or struct")
	}

	//
	//

	cols := make([]int, rstructt.NumField())
	for i := 0; i < len(cols); i++ {
		col, err := abbrev(rstructt.Field(i).Name, file.keys...)
		if err != nil {
			return &Error{"Unmarshal", "abbreviation", err}
		}
		cols[i] = col
	}

	if !IsMap && (IsSlice || IsStruct) {
		rsl := reflect.New(rslt).Elem()
		flds := map[int]byte{}
		rfirst := reflect.New(rstructt).Elem()

		for _, ln := range file.body {
			rstruct := reflect.New(rstructt).Elem()

			for f, col := range cols {
				rfld := rstruct.Field(f)
				rfld2, err := parse(rfld.Type(), ln[col])
				if err != nil {
					return &Error{"Unmarshal", "parsing data", err}
				}
				if IsStruct {
					if len(flds) == len(cols) {
						rv.Set(rfirst)
						return nil
					}
					if _, found := flds[f]; len(ln[col]) == 0 || found {
						continue
					}
					rfirst.Field(f).Set(rfld2)
					flds[f] = 0
				} else {
					rfld.Set(rfld2)
				}
			}

			if IsSlice {
				rsl = reflect.Append(rsl, rstruct)
			}
		}

		rv.Set(rsl)
		return nil
	}

	if IsMap {
		rkeyt := rt.Key()
		rmap := reflect.MakeMap(rt)

		key := rslt.Name()
		if len(key) == 0 {
			key = rstructt.Name()
		}
		keyCol, err := abbrev(key, file.keys...)
		if err != nil {
			return &Error{"Unmarshal", "abbreviation", err}
		}

		for _, ln := range file.body {
			if keyCol >= len(ln) {
				continue
			}

			rstruct := reflect.New(rstructt).Elem()

			key := ln[keyCol]
			rkey, err := parse(rkeyt, key)
			if err != nil {
				return &Error{"Unmarshal", "parsing key", err}
			}

			for f, col := range cols {
				rfld := rstruct.Field(f)
				rfld2, err := parse(rfld.Type(), ln[col])
				if err != nil {
					return &Error{"Unmarshal", "parsing data", err}
				}
				rfld.Set(rfld2)
			}

			rmapsl2 := rmap.MapIndex(rkey)
			if rmapsl2.Kind() == reflect.Invalid {
				rmapsl2 = reflect.New(rslt).Elem()
			}
			rmap.SetMapIndex(rkey, reflect.Append(rmapsl2, rstruct))
		}

		rv.Set(rmap)
	}

	return nil
}

// Add adds lines to the file's underlying body.
func (file *File) Add(lines ...[]string) {
	file.body = append(file.body, lines...)
}

// Save saves the Excel file.
func (file *File) Save(name string) error {
	if err := file.init(); err != nil {
		return &Error{"Save", "init", err}
	}

	f, err := os.Create(name)
	if err != nil {
		return &Error{"Save", "creating file", err}
	}
	defer f.Close()

	w := csv.NewWriter(f)
	w.Comma = file.Comma
	if err := w.WriteAll(append(append([][]string{}, file.keys), file.body...)); err != nil {
		return &Error{"Save", "writing data", err}
	}

	return nil
}

func parse(rt reflect.Type, v string) (reflect.Value, error) {
	v = strings.Trim(v, " ")
	rv := reflect.New(rt).Elem()
	switch rv.Kind() {
	case reflect.String:
		rv.Set(reflect.ValueOf(v).Convert(rt))

	case reflect.Float64:
		v = strings.Replace(v, ",", "", -1)
		x, err := strconv.ParseFloat(v, 64)
		if err != nil && len(v) > 0 {
			x, err = clean(v)
			if err != nil {
				return rv, err
			}
		}
		rv.Set(reflect.ValueOf(x).Convert(rt))

	case reflect.Int:
		v = strings.Replace(v, ",", "", -1)
		x, err := strconv.Atoi(v)
		if err != nil && len(v) > 0 {
			return rv, err
		}
		rv.Set(reflect.ValueOf(x).Convert(rt))

	case reflect.ValueOf(time.Time{}).Kind():
		var x time.Time
		for _, layout := range layouts {
			if t, err := time.Parse(layout, v); err == nil {
				x = t
				break
			}
		}
		if x == (time.Time{}) && len(v) > 0 {
			return rv, errors.New("bad time format '" + v + "'")
		}
		rv.Set(reflect.ValueOf(x).Convert(rt))
	}

	return rv, nil
}

func clean(s string) (float64, error) {
	mess, err := regexp.Compile("[0-9.,]+")
	if err != nil {
		return 0, err
	}
	amt, err := strconv.ParseFloat(mess.FindString(s), 64)
	if err != nil {
		return 0, err
	}
	return amt / 100, nil
}

func abbrev(sub string, strs ...string) (int, error) {
	if len(sub) == 0 {
		return -1, errors.New("empty substring")
	}
	expr := `(?i)[^` + string(sub[0]) + `]*`
	for x, r := range sub {
		end := `.*`
		if xplus1 := x + 1; xplus1 < len(sub) {
			end = `[^` + string(sub[xplus1]) + `]*`
		}
		expr += string(r) + end
	}

	n := -1
	matches := []string{""}
	for i, str := range strs {
		l, L := len(str), len(matches[0])
		if l == 0 || !regexp.MustCompile(expr).MatchString(str) {
			continue
		} else if 0 < L && L < l {
			continue
		}

		if l == L {
			matches = append(matches, str)
		} else {
			matches = []string{str}
			n = i
		}
	}

	if len(matches[0]) == 0 {
		return -1, errors.New("no match for '" + sub + "' in {" + strings.Join(strs, ", ") + "}")
	} else if len(matches) > 1 {
		return -1, errors.New(sub + "=" + fmt.Sprint(matches) + " (too many matches)")
	}
	return n, nil
}

// regef finds a file matching the expression. If not found,
// it finds a match for the assumed regular expression.
func regef(expr string) (string, error) {
	if _, err := os.Stat(expr); err != nil {
		ex, err2 := regexp.Compile(`(?i)` + expr)
		if err2 != nil {
			return "", err
		}
		dir := ""
		if i := strings.LastIndex(expr, "/"); i > -1 {
			dir = expr[:i]
		}
		var matches []string
		if err := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
			path = strings.Replace(path, `\`, "/", -1)
			if ex.MatchString(path) {
				matches = append(matches, path)
			}
			return nil
		}); err != nil {
			return "", err
		}
		if len(matches) == 0 {
			return "", errors.New("no matches")
		}
		sort.Strings(matches)
		return matches[0], nil
	}
	return expr, nil
}
