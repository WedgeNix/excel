# xl
Fuzzy xlsx/xls/csv decoder

```go
// filenames: "this_is_a_test (1).csv"
//            "this_is_a_test (2).xlsx"
//            "ThisTest.xlsx"

file, _ := xl.Open(`_is_a_`, xl.OpenLast)

// the following struct tags, 'key' and 'value',
// allow regular expression matches for headers (keys) and columns (values)

var data []struct {
    SKU string `key:"[Cc]atalog(ue)? ?[Nn]o" value:"[^-]{4}-[0-9]{4}"`
    ATS int    `key:"O[Tt][Ss]|A[Tt][Ss]"`
}
file.Decode(&data)
```
