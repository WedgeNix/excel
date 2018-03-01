# excel
Unmarshaler for xlsx/xls/csv

```go
// col1 , col2 , date
//  val ,  2.1 , 2018/3/1
//  val ,  0.9 , 2018/2/1
//  abc ,    2 , 2018/1/1

f := excel.File{Name: "test.csv"}

type Date struct {
    Col1 string
    Col2 float64
}
var dates map[time.Time][]Date
if err := f.Unmarshal(&dates); err != nil {
    // handle err
}

for date, ln := range dates {
    //
}

//
//

f := excel.File{Name: "test.csv"}

type whatever struct {
    Col1 string
    Col2 float64
    Date time.Time
}
var data []whatever
if err := f.Unmarshal(&data); err != nil {
    // handle err
}

for _, ln := range data {
    //
}
```
