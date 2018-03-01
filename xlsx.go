package excel

import "github.com/tealeg/xlsx"

type xlsxf struct{ x *xlsx.File }

func (x xlsxf) Sheets() int {
	return len(x.x.Sheets)
}
func (x xlsxf) Rows(sheet int) int {
	return len(x.x.Sheets[sheet].Rows)
}
func (x xlsxf) Cols(sheet, row int) int {
	return len(x.x.Sheets[sheet].Rows[row].Cells)
}
func (x xlsxf) Cell(sheet, row, col int) string {
	return x.x.Sheets[sheet].Rows[row].Cells[col].String()
}
