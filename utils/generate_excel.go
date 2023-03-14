package utils

import (
	"encoding/csv"
	"fmt"
	"os"

	"github.com/kwamekyeimonies/Golang-Excelization/component"
	"github.com/xuri/excelize/v2"
)

const (
	Sheetname = "Expense Report"
)

func Generate_CSV(f *excelize.File, start, end  component.Axis) error {
	var data [][]string

	for i := start.Row; i <= end.Row; i++ {
		row := []string{}
		for j := []rune(start.Col)[0]; j <= []rune(end.Col)[0]; j++ {
			value, err := f.GetCellValue(Sheetname, fmt.Sprintf("%s%d", string(j), i), excelize.Options{})
			if err != nil {
				return err
			}
			row = append(row, value)
		}
		data = append(data, row)
	}

	file, err := os.Create("expenses.csv")
	if err != nil {
		return err
	}
	defer f.Close()

	writer := csv.NewWriter(file)
	return writer.WriteAll(data)
}
