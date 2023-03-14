package main

import (
	"fmt"
	"log"

	"github.com/kwamekyeimonies/Golang-Excelization/component"
	"github.com/kwamekyeimonies/Golang-Excelization/utils"
	"github.com/xuri/excelize/v2"
)

const (
	Sheetname = "Expense Report"
)

type Axis struct {
	row int
	col string
}

var (
	expenseData = [][]interface{}{
		{"2022-04-10", "Flight", "Trip to Dubai", "", "", "$3,462.00"},
		{"2022-04-10", "Hotel", "Trip to Cannada", "", "", "$1,280.00"},
		{"2022-04-12", "Swags", "App launch", "", "", "$862.00"},
		{"2022-03-15", "Marketing", "Software Test", "", "", "$7,520.00"},
		{"2022-04-11", "Event hall", "App launch", "", "", "$2,080.00"},
	}
)

func main() {
	
		start := component.Axis{ Row:1, Col: "A"}
		end := component.Axis{Row: 10,  Col: "B"}

		err = f.SaveAs("expense-report.xlsx")
		if err !=nil{
			log.Fatal(err)
		}
		err = utils.Generate_CSV(f,start,end)

		if err != nil {
			log.Fatal(err)
		}
	

}
