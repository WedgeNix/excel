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

var excel = File{
	Comma: ',',
}

type Err struct{ error }

func (e Err) Error() string {
	return "excel: " + error(e).Error()
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

		var sheets [][][]string

		ext := filepath.Ext(file.Name)
		switch ext {
		case ".xlsx":
			c, err := xlsx.FileToSlice(file.Name)
			if err != nil {
				initErr = err
				return
			}
			sheets = c
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
			f.Close()
			sheets = [][][]string{c}
		case ".xls":
			f, err := xls.Open(file.Name, "")
			if err != nil {
				initErr = err
				return
			}
			sheets = [][][]string{f.ReadAllCells()}
		default:
			initErr = errors.New("'" + ext + "' is not an Excel format")
			return
		}

		//
		//

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
				if len(cells) > longest {
					longest = len(cells)
					keys = cells
					body = rows[i+1:]
				}
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

// Unmarshal parses the excel-encoded data and stores the result
// in the value pointed to by ptr.
func (file *File) Unmarshal(ptr interface{}, opt ...interface{}) error {
	if err := file.init(); err != nil {
		return Err{err}
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
			return Err{err}
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
					return Err{err}
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
			return Err{err}
		}

		for _, ln := range file.body {
			if keyCol >= len(ln) {
				continue
			}

			rstruct := reflect.New(rstructt).Elem()

			key := ln[keyCol]
			rkey, err := parse(rkeyt, key)
			if err != nil {
				return Err{errors.New("could not parse time '" + key + "'")}
			}

			for f, col := range cols {
				rfld := rstruct.Field(f)
				rfld2, err := parse(rfld.Type(), ln[col])
				if err != nil {
					return Err{err}
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
		return Err{err}
	}

	f, err := os.Create(name)
	if err != nil {
		return Err{err}
	}
	defer f.Close()

	w := csv.NewWriter(f)
	w.Comma = file.Comma
	return Err{w.WriteAll(append(append([][]string{}, file.keys), file.body...))}
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
			return rv, errors.New("bad time format '" + v + "'")
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
		return -1, errors.New("no match for '" + sub + "' in {" + strings.Join(strs, ", ") + "}")
	} else if len(matches) > 1 {
		return -1, errors.New(sub + "=" + fmt.Sprint(matches) + " (too many matches)")
	}
	return n, nil
}
