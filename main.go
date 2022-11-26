package main

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

func main() {
	header := map[string]string{
		"A1": "No", "B1": "Name", "C1": "Age", "D1": "Alamat",
	}

	body := map[string]string{
		"A2": "1", "B2": "Andi", "C2": "27", "D2": "Bogor",
		"A3": "2", "B3": "Nana", "C3": "26", "D3": "Bogor",
	}

	// initialize new file
	f := excelize.NewFile()

	// Set value of a cell.
	for k, v := range header {
		f.SetCellValue("Sheet1", k, v)
	}
	for k, v := range body {
		f.SetCellValue("Sheet1", k, v)
	}

	// give style
	style, err := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "gradient", Color: []string{"#9BE3DE", "#BEEBE9"}, Shading: 1},
		Font: &excelize.Font{
			Bold:   true,
			Italic: false,
			Size:   12,
		},
	})
	if err != nil {
		fmt.Println(err)
	}
	err = f.SetCellStyle("Sheet1", "A1", "D1", style)

	// Save spreadsheet by the given path.
	if err := f.SaveAs("./excel/Book1.xlsx"); err != nil {
		fmt.Println(err)
	}
}
