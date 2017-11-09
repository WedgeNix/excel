package xl

import "github.com/extrame/xls"

type xlsf struct{ x *xls.WorkBook }

func (x xlsf) Sheets() int {
	return x.x.NumSheets()
}

func (x xlsf) Rows(sheet int) int {
	return int(x.x.GetSheet(sheet).MaxRow) + 1
}

func (x xlsf) Cols(sheet, row int) int {
	return x.x.GetSheet(sheet).Row(row).LastCol() + 1
}

func (x xlsf) Cell(sheet, row, col int) string {
	return x.x.GetSheet(sheet).Row(row).Col(col)
}
