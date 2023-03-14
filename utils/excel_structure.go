package utils

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func Excel_Structuring() {
	var err error
	f := excelize.NewFile()
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		log.Fatal(err.Error())
	}
	f.SetActiveSheet(index)
	f.SetSheetName("Sheet1", Sheetname)

	err = f.SetRowHeight(Sheetname, 1, 12)
	err = f.MergeCell(Sheetname, "A1", "H1")

	err = f.SetRowHeight(Sheetname, 2, 25)
	err = f.MergeCell(Sheetname, "B2", "D2")

	style, err := f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 20, Color: "6d64e8"}})
	err = f.SetCellStyle(Sheetname, "B2", "D2", style)
	err = f.SetSheetRow(Sheetname, "B2", &[]interface{}{"Tea-Code Dev."})
	err = f.MergeCell(Sheetname, "B3", "D3")
	err = f.SetSheetRow(Sheetname, "B3", &[]interface{}{"Accra Ghana"})

	err = f.MergeCell(Sheetname, "B4", "D4")
	err = f.SetSheetRow(Sheetname, "B4", &[]interface{}{"Indianapolis, IN 46276"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.MergeCell(Sheetname, "B5", "D5")
	err = f.SetCellStyle(Sheetname, "B5", "D5", style)
	err = f.SetSheetRow(Sheetname, "B5", &[]interface{}{"(233) 558-485290"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 32, Color: "2B4492", Bold: true}})
	err = f.MergeCell(Sheetname, "B7", "G7")
	err = f.SetCellStyle(Sheetname, "B7", "G7", style)
	err = f.SetSheetRow(Sheetname, "B7", &[]interface{}{"Expense Report"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Color: "E25184", Bold: true}})
	err = f.MergeCell(Sheetname, "B8", "C8")
	err = f.SetCellStyle(Sheetname, "B8", "C8", style)
	err = f.SetSheetRow(Sheetname, "B8", &[]interface{}{"09/04/00 - 09/05/00"})

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	err = f.SetCellStyle(Sheetname, "B10", "G10", style)
	err = f.SetSheetRow(Sheetname, "B10", &[]interface{}{"Name", "", "Employee ID", "", "Department"})
	err = f.MergeCell(Sheetname, "B10", "C10")
	err = f.MergeCell(Sheetname, "D10", "E10")
	err = f.MergeCell(Sheetname, "F10", "G10")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.SetCellStyle(Sheetname, "B11", "G11", style)
	err = f.SetSheetRow(Sheetname, "B11", &[]interface{}{"Tenkorang Daniel", "", "#1B800XR", "", "Software Engineering"})
	err = f.MergeCell(Sheetname, "B11", "C11")
	err = f.MergeCell(Sheetname, "D11", "E11")
	err = f.MergeCell(Sheetname, "F11", "G11")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	err = f.SetCellStyle(Sheetname, "B13", "G13", style)
	err = f.SetSheetRow(Sheetname, "B13", &[]interface{}{"Manager", "", "Purpose"})
	err = f.MergeCell(Sheetname, "B13", "C13")
	err = f.MergeCell(Sheetname, "D13", "E13")

	style, err = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	err = f.SetCellStyle(Sheetname, "B14", "G14", style)
	err = f.SetSheetRow(Sheetname, "B14", &[]interface{}{"Jane Doe", "", "Brand Campaign"})
	err = f.MergeCell(Sheetname, "B14", "C14")
	err = f.MergeCell(Sheetname, "D14", "E14")

	style, err = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "2B4492"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	err = f.SetCellStyle(Sheetname, "B17", "G17", style)
	err = f.SetSheetRow(Sheetname, "B17", &[]interface{}{"Date", "Category", "Description", "", "Notes", "Amount"})
	err = f.MergeCell(Sheetname, "D17", "E17")
	err = f.SetRowHeight(Sheetname, 17, 32)

	startRow := 18
	for i := startRow; i < (len(expenseData) + startRow); i++ {
		var fill string
		if i%2 == 0 {
			fill = "F3F3F3"
		} else {
			fill = "FFFFFF"
		}

		style, err = f.NewStyle(&excelize.Style{
			Fill:      excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{fill}},
			Font:      &excelize.Font{Color: "666666"},
			Alignment: &excelize.Alignment{Vertical: "center"},
		})
		err = f.SetCellStyle(Sheetname, fmt.Sprintf("B%d", i), fmt.Sprintf("G%d", i), style)
		err = f.SetSheetRow(Sheetname, fmt.Sprintf("B%d", i), &expenseData[i-18])
		err := f.SetCellRichText(Sheetname, fmt.Sprintf("C%d", i), []excelize.RichTextRun{
			{Text: expenseData[i-18][1].(string), Font: &excelize.Font{Bold: true}},
		})

		if err != nil {
			log.Fatal(err.Error())
		}

		err = f.MergeCell(Sheetname, fmt.Sprintf("D%d", i), fmt.Sprintf("E%d", i))
		err = f.SetRowHeight(Sheetname, i, 18)
	}

	err = f.SaveAs("expense-report.xlsx")
	if err != nil {
		log.Fatal(err)
	}

}
