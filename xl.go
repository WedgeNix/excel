package xl

import (
	"encoding/csv"
	"errors"
	"log"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"sort"
	"strconv"
	"strings"

	"github.com/extrame/xls"

	"github.com/tealeg/xlsx"
)

type File struct {
	xlsx *xlsx.File
	xls  *xls.WorkBook
	csv  [][]string
}

type OpenFunc func(a, b string) bool

var (
	OpenFirst = func(a, b string) bool { return a < b }
	OpenLast  = func(a, b string) bool { return a > b }
)

func Open(expr string, better OpenFunc) (*File, error) {
	r, err := regexp.Compile(expr)
	if err != nil {
		return nil, err
	}

	var matches []string
	if filepath.Walk(".", func(path string, info os.FileInfo, err error) error {
		name := info.Name()
		// println("checking '" + path + "' for '" + name + "'")
		if r.MatchString(name) {
			matches = append(matches, name)
		}
		return err
	}); err != nil {
		return nil, err
	}
	if len(matches) == 0 {
		return nil, errors.New("no match")
	}

	log.Println(matches)

	if better == nil && len(matches) > 1 {
		return nil, errors.New("too many matches")
	} else if better != nil {
		sort.Slice(matches, func(i, j int) bool {
			return better(matches[i], matches[j])
		})
	}
	match := matches[0]
	log.Println([]string{match})

	var xlf File
	if i := strings.LastIndex(match, "."); i != -1 && match[i:] == ".xlsx" {
		println(`found ` + match[i:])
		xlf.xlsx, err = xlsx.OpenFile(match)
	} else if match[i:] == ".csv" {
		println(`found ` + match[i:])
		f, err := os.Open(match)
		if err != nil {
			return nil, err
		}
		r := csv.NewReader(f)
		xlf.csv, err = r.ReadAll()
		if err != nil {
			return nil, err
		}
		err = f.Close()
	} else if match[i:] == ".xls" {
		xlf.xls, err = xls.Open(match, "")
	}
	if err != nil {
		return nil, err
	}

	return &xlf, nil
}

type heuristics struct {
	index     int
	kind      reflect.Kind
	k, v      *regexp.Regexp
	shColErrs [][3]int
}

func (f File) Decode(v interface{}) error {
	ptr := reflect.ValueOf(v)
	if ptr.Kind() != reflect.Ptr {
		return errors.New("not a pointer")
	}

	val := reflect.Indirect(ptr)
	t := val.Type()
	if t.Kind() != reflect.Slice {
		return errors.New("not a slice")
	}

	e := t.Elem()
	if e.Kind() != reflect.Struct {
		return errors.New("not a slice of structs")
	}

	var heurs []*heuristics

	for i := 0; i < e.NumField(); i++ {
		fld := e.Field(i)
		key, ok := fld.Tag.Lookup("key")
		if !ok {
			return errors.New("missing struct tag 'key'")
		}
		kregex, err := regexp.Compile(key)
		if err != nil {
			return err
		}
		value := fld.Tag.Get("value")
		vregex, err := regexp.Compile(value)
		if err != nil {
			return err
		}
		if len(value) == 0 {
			vregex = nil
		}
		heurs = append(heurs, &heuristics{
			index: i,
			kind:  fld.Type.Kind(),
			k:     kregex,
			v:     vregex,
		})
	}

	if f.xls != nil {
		for Si := 0; Si < f.xls.NumSheets(); Si++ {
			s := f.xls.GetSheet(Si)
			if s.MaxRow < 1 {
				continue
			}

			// determine columns using first row
			r := s.Row(0)
			for Ci := 0; Ci <= r.LastCol(); Ci++ {
				c := r.Col(Ci)
				for _, H := range heurs {
					if !H.k.MatchString(c) {
						continue
					}
					H.shColErrs = append(H.shColErrs, [3]int{Si, Ci, 0})
					// log.Println(`[3]int{Si, Ci, 0}`, [3]int{Si, Ci, 0})
				}
			}
		}

		var maxL int

		// check all matches at (sheet, column) and count parsing errors
		for _, H := range heurs {
			for sce, shColErrs := range H.shColErrs {
				Si, Ci, errs := shColErrs[0], shColErrs[1], shColErrs[2]
				s := f.xls.GetSheet(Si)
				maxL = int(s.MaxRow) + 1

				for Ri := 1; Ri < maxL; Ri++ {
					c := s.Row(Ri).Col(Ci)

					if H.v != nil && !H.v.MatchString(c) {
						errs++
						continue
					}

					switch H.kind {
					case reflect.Int:
						_, err := strconv.Atoi(c)
						if err != nil {
							errs++
						}
					case reflect.Float64:
						_, err := strconv.ParseFloat(c, 64)
						if err != nil {
							errs++
						}
					}
				}
				H.shColErrs[sce][2] = errs
			}
		}
		// println(`maxL`, maxL)

		val.Set(reflect.MakeSlice(t, maxL-1, maxL-1))

		for _, H := range heurs {
			sort.Slice(H.shColErrs, func(i, j int) bool {
				return H.shColErrs[i][2] < H.shColErrs[j][2]
			})
			if H.shColErrs[0][2] == maxL {
				return errors.New("no pattern match")
			}
			Si := H.shColErrs[0][0]
			Ci := H.shColErrs[0][1]
			i := H.index

			for Ri := 1; Ri < maxL; Ri++ {
				c := f.xls.GetSheet(Si).Row(Ri).Col(Ci)

				var rv reflect.Value
				switch H.kind {
				case reflect.Int:
					x, _ := strconv.Atoi(c)
					rv = reflect.ValueOf(x)
				case reflect.Float64:
					x, _ := strconv.ParseFloat(c, 64)
					rv = reflect.ValueOf(x)
				case reflect.String:
					rv = reflect.ValueOf(c)
				default:
					return errors.New("type not supported")
				}

				val.Index(Ri - 1).Field(i).Set(rv)
			}

			// log.Printf("reflectIdx:%d sheet:%d col:%d", i, Si, Ci)
		}
	}

	return nil
}

// func (f File) decodeXLSX(v interface{}) error {

// }

// func (f File) decodeXLS(v interface{}) error {

// }

// func (f File) decodeCSV(v interface{}) error {

// }
