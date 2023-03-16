package main

import (
	"fmt"
	"log"

	"github.com/kwamekyeimonies/Golang-Excelization/component"
	"github.com/kwamekyeimonies/Golang-Excelization/utils"
	"github.com/xuri/excelize/v2"
)

var (
	expenseData = [][]interface{}{
		{"2022-04-10", "Flight", "Trip to Dubai", "", "", "$3,462.00"},
		{"2022-04-10", "Hotel", "Trip to Cannada", "", "", "$1,280.00"},
		{"2022-04-12", "Swags", "App launch", "", "", "$862.00"},
		{"2022-03-15", "Marketing", "Software Test", "", "", "$7,520.00"},
		{"2022-04-11", "Event hall", "App launch", "", "", "$2,080.00"},
	}
)

const (
	Sheetname = "Expense Report"
)

func main() {

	f := excelize.NewFile()
	index, err := f.NewSheet("Sheet1")
	if err != nil {
		log.Fatal(err.Error())
	}
	f.SetActiveSheet(index)
	f.SetSheetName("Sheet1", Sheetname)

	_ = f.SetRowHeight(Sheetname, 1, 12)
	_ = f.MergeCell(Sheetname, "A1", "H1")

	_ = f.SetRowHeight(Sheetname, 2, 25)
	_ = f.MergeCell(Sheetname, "B2", "D2")

	style, _ := f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 20, Color: "6d64e8"}})
	_ = f.SetCellStyle(Sheetname, "B2", "D2", style)
	_ = f.SetSheetRow(Sheetname, "B2", &[]interface{}{"Tea-Code Dev."})
	_ = f.MergeCell(Sheetname, "B3", "D3")
	_ = f.SetSheetRow(Sheetname, "B3", &[]interface{}{"Accra Ghana"})

	_ = f.MergeCell(Sheetname, "B4", "D4")
	_ = f.SetSheetRow(Sheetname, "B4", &[]interface{}{"Indianapolis, IN 46276"})

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	_ = f.MergeCell(Sheetname, "B5", "D5")
	_ = f.SetCellStyle(Sheetname, "B5", "D5", style)
	_ = f.SetSheetRow(Sheetname, "B5", &[]interface{}{"(233) 558-485290"})

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 32, Color: "2B4492", Bold: true}})
	_ = f.MergeCell(Sheetname, "B7", "G7")
	_ = f.SetCellStyle(Sheetname, "B7", "G7", style)
	_ = f.SetSheetRow(Sheetname, "B7", &[]interface{}{"Expense Report"})

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Color: "E25184", Bold: true}})
	_ = f.MergeCell(Sheetname, "B8", "C8")
	_ = f.SetCellStyle(Sheetname, "B8", "C8", style)
	_ = f.SetSheetRow(Sheetname, "B8", &[]interface{}{"09/04/00 - 09/05/00"})

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	_ = f.SetCellStyle(Sheetname, "B10", "G10", style)
	_ = f.SetSheetRow(Sheetname, "B10", &[]interface{}{"Name", "", "Employee ID", "", "Department"})
	_ = f.MergeCell(Sheetname, "B10", "C10")
	_ = f.MergeCell(Sheetname, "D10", "E10")
	_ = f.MergeCell(Sheetname, "F10", "G10")

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	_ = f.SetCellStyle(Sheetname, "B11", "G11", style)
	_ = f.SetSheetRow(Sheetname, "B11", &[]interface{}{"Tenkorang Daniel", "", "#1B800XR", "", "Software Engineering"})
	_ = f.MergeCell(Sheetname, "B11", "C11")
	_ = f.MergeCell(Sheetname, "D11", "E11")
	_ = f.MergeCell(Sheetname, "F11", "G11")

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Size: 13, Bold: true}})
	_ = f.SetCellStyle(Sheetname, "B13", "G13", style)
	_ = f.SetSheetRow(Sheetname, "B13", &[]interface{}{"Manager", "", "Purpose"})
	_ = f.MergeCell(Sheetname, "B13", "C13")
	_ = f.MergeCell(Sheetname, "D13", "E13")

	style, _ = f.NewStyle(&excelize.Style{Font: &excelize.Font{Color: "666666"}})
	_ = f.SetCellStyle(Sheetname, "B14", "G14", style)
	_ = f.SetSheetRow(Sheetname, "B14", &[]interface{}{"Jane Doe", "", "Brand Campaign"})
	_ = f.MergeCell(Sheetname, "B14", "C14")
	_ = f.MergeCell(Sheetname, "D14", "E14")

	style, _ = f.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Size: 13, Bold: true, Color: "2B4492"},
		Alignment: &excelize.Alignment{Vertical: "center"},
	})
	_ = f.SetCellStyle(Sheetname, "B17", "G17", style)
	_ = f.SetSheetRow(Sheetname, "B17", &[]interface{}{"Date", "Category", "Description", "", "Notes", "Amount"})
	_ = f.MergeCell(Sheetname, "D17", "E17")
	_ = f.SetRowHeight(Sheetname, 17, 32)

	startRow := 18
	for i := startRow; i < (len(expenseData) + startRow); i++ {
		var fill string
		if i%2 == 0 {
			fill = "F3F3F3"
		} else {
			fill = "FFFFFF"
		}

		style, _ = f.NewStyle(&excelize.Style{
			Fill:      excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{fill}},
			Font:      &excelize.Font{Color: "666666"},
			Alignment: &excelize.Alignment{Vertical: "center"},
		})
		_ = f.SetCellStyle(Sheetname, fmt.Sprintf("B%d", i), fmt.Sprintf("G%d", i), style)
		_ = f.SetSheetRow(Sheetname, fmt.Sprintf("B%d", i), &expenseData[i-18])
		err := f.SetCellRichText(Sheetname, fmt.Sprintf("C%d", i), []excelize.RichTextRun{
			{Text: expenseData[i-18][1].(string), Font: &excelize.Font{Bold: true}},
		})

		if err != nil {
			log.Fatal(err.Error())
		}

		_ = f.MergeCell(Sheetname, fmt.Sprintf("D%d", i), fmt.Sprintf("E%d", i))
		_ = f.SetRowHeight(Sheetname, i, 18)
	}

	err = f.SaveAs("expense-report.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	start := component.Axis{Row: 1, Col: "A"}
	end := component.Axis{Row: 10, Col: "B"}

	utils.Excel_Structuring()
	errors := utils.Generate_CSV(f, start, end)
	if errors != nil {
		log.Fatal(err)
	}

}
