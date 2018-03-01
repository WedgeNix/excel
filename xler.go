package excel

type xler interface {
	Sheets() int
	Rows(int) int
	Cols(int, int) int
	Cell(int, int, int) string
}
