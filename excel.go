package excel

import (
	"encoding/csv"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/extrame/xls"
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
	"2006/1/_2",
	"2006/01/02",
	"1/_2/2006",
	"01/02/2006",
	"2006-1-_2",
	"2006-01-02",
	"1-_2-2006",
	"01-02-2006",
}

var excel = File{
	Comma: ',',
}

// File is the generic Excel API controller.
type File struct {
	initOnce sync.Once

	Name  string
	Comma rune

	keys []string
	body [][]string
}

func (file *File) init() error {
	var initErr error
	file.initOnce.Do(func() {
		if file.Comma == 0 {
			file.Comma = excel.Comma
		}

		var x xler
		ext := filepath.Ext(file.Name)
		switch ext {
		case ".xlsx":
			f, err := xlsx.OpenFile(file.Name)
			if err != nil {
				initErr = err
				return
			}
			x = xlsxf{f}
		case ".csv", ".txt":
			f, err := os.Open(file.Name)
			if err != nil {
				initErr = err
				return
			}
			r := csv.NewReader(f)
			r.Comma = file.Comma
			r.LazyQuotes = true
			c, err := r.ReadAll()
			if err != nil {
				initErr = err
				return
			}
			err = f.Close()
			x = csvf{c}
		case ".xls":
			f, err := xls.Open(file.Name, "")
			if err != nil {
				initErr = err
				return
			}
			x = xlsf{f}
		default:
			initErr = errors.New("'" + ext + "' is not an Excel format")
			return
		}

		//
		//

		var keys []string
		var body [][]string
		sheetCnt := x.Sheets()
		for sheeti := 0; sheeti < sheetCnt; sheeti++ {
			rowCnt := x.Rows(sheeti)
			for rowi := 0; rowi < rowCnt; rowi++ {
				ln := make([]string, x.Cols(sheeti, rowi))
				for coli := 0; coli < len(ln); coli++ {
					ln[coli] = x.Cell(sheeti, rowi, coli)
				}
				if keys == nil {
					keys = ln
					continue
				}
				body = append(body, ln)
			}
		}

		switch {
		case len(keys) == 0:
			initErr = errors.New("empty file")
		case len(body) == sheetCnt:
			initErr = errors.New("no values in file")
		default:
			file.keys = keys
			file.body = body
		}
	})
	return initErr
}

// Unmarshal parses the excel-encoded data and stores the result
// in the value pointed to by ptr.
func (file *File) Unmarshal(ptr interface{}, opt ...interface{}) error {
	if err := file.init(); err != nil {
		return err
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

	var IsMap, IsSlice bool
	rslt := rt
	switch rt.Kind() {
	case reflect.Map:
		IsMap = true
		rslt = rv.Type().Elem()
	case reflect.Slice:
		IsSlice = true
	default:
		panic("not a map or slice")
	}

	if rslt.Kind() != reflect.Slice {
		panic("not a slice map")
	}
	rstructt := rslt.Elem()
	if rstructt.Kind() != reflect.Struct {
		panic("not a struct slice map")
	}

	//
	//

	cols := make([]int, rstructt.NumField())
	for i := 0; i < len(cols); i++ {
		col, err := abbrev(rstructt.Field(i).Name, file.keys...)
		if err != nil {
			return err
		}
		cols[i] = col
	}

	if IsSlice {
		rsl := reflect.New(rslt).Elem()

		for _, ln := range file.body {
			rstruct := reflect.New(rstructt).Elem()

			for f, col := range cols {
				rfld := rstruct.Field(f)
				rfld2, err := parse(rfld.Type(), ln[col])
				if err != nil {
					return err
				}
				rfld.Set(rfld2)
			}

			rsl = reflect.Append(rsl, rstruct)
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
			return err
		}

		for _, ln := range file.body {
			rstruct := reflect.New(rstructt).Elem()

			key := ln[keyCol]
			rkey, err := parse(rkeyt, key)
			if err != nil {
				return err
			}

			for f, col := range cols {
				rfld := rstruct.Field(f)
				rfld2, err := parse(rfld.Type(), ln[col])
				if err != nil {
					return err
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

// Save saves the Excel file.
func (file *File) Save(name string) error {
	if err := file.init(); err != nil {
		return err
	}

	f, err := os.Create(name)
	if err != nil {
		return err
	}
	defer f.Close()

	w := csv.NewWriter(f)
	w.Comma = file.Comma
	return w.WriteAll(append(append([][]string{}, file.keys), file.body...))
}

func parse(rt reflect.Type, v string) (reflect.Value, error) {
	rv := reflect.New(rt).Elem()
	switch rv.Kind() {
	case reflect.String:
		rv.Set(reflect.ValueOf(v))

	case reflect.Float64:
		v = strings.Replace(v, ",", "", -1)
		x, err := strconv.ParseFloat(v, 64)
		if err != nil && len(v) > 0 {
			return rv, err
		}
		rv.Set(reflect.ValueOf(x))

	case reflect.Int:
		v = strings.Replace(v, ",", "", -1)
		x, err := strconv.Atoi(v)
		if err != nil && len(v) > 0 {
			return rv, err
		}
		rv.Set(reflect.ValueOf(x))

	case reflect.ValueOf(time.Time{}).Kind():
		var x time.Time
		for _, layout := range layouts {
			if t, err := time.Parse(layout, v); err == nil {
				x = t
				break
			}
		}
		if x == (time.Time{}) && len(v) > 0 {
			return rv, errors.New("bad time format")
		}
		rv.Set(reflect.ValueOf(x))
	}

	return rv, nil
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
		return -1, errors.New("no match for '" + sub + "'")
	} else if len(matches) > 1 {
		return -1, errors.New(sub + "=" + fmt.Sprint(matches) + " (too many matches)")
	}
	return n, nil
}
