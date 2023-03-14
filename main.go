package main

import (
	"log"

	"github.com/kwamekyeimonies/Golang-Excelization/component"
	"github.com/kwamekyeimonies/Golang-Excelization/utils"
	"github.com/xuri/excelize/v2"
)

const (
	Sheetname = "Expense Report"
)

func main() {

	f := excelize.NewFile()
	start := component.Axis{Row: 1, Col: "A"}
	end := component.Axis{Row: 10, Col: "B"}

	utils.Excel_Structuring()
	err := utils.Generate_CSV(f, start, end)
	if err != nil {
		log.Fatal(err)
	}

}
