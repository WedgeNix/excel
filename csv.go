package xl

type csvf struct{ c [][]string }

func (c csvf) Sheets() int {
	return 1
}

func (c csvf) Rows(sheet int) int {
	return len(c.c)
}

func (c csvf) Cols(sheet, row int) int {
	return len(c.c[row])
}

func (c csvf) Cell(sheet, row, col int) string {
	if L := c.Cols(sheet, row); col >= L {
		println("cols", col, "vs", L-1)
	}
	return c.c[row][col]
}
