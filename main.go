package main

import (
	"fmt"
	"strconv"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	xlsx2 "github.com/tealeg/xlsx"
)

func main() {
	xls1, err := excelize.OpenFile("resources/turnover.xlsx")
	if err != nil {
		fmt.Println("oops:", err)
		return
	}
	xls2, err := xlsx2.OpenFile("resources/turnover.xlsx")
	if err != nil {
		fmt.Println("oops:", err)
		return
	}
	start := time.Now()
	xls1Save(xls1)
	fmt.Println("finish xls1: ", time.Since(start))
	start = time.Now()
	xls2Save(xls2)
	fmt.Println("finish xls2: ", time.Since(start))
}

func xls1Save(xlsx *excelize.File) {
	sheet := xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
	for i := 0; i < 15000; i++ {
		N := strconv.Itoa(i + 1)
		xlsx.SetCellStr(sheet, "A"+N, "Hello World")
		xlsx.SetCellStr(sheet, "B"+N, "Hello World")
		xlsx.SetCellStr(sheet, "C"+N, "Hello World")
		xlsx.SetCellStr(sheet, "D"+N, "Hello World")
		xlsx.SetCellStr(sheet, "E"+N, "Hello World")
		xlsx.SetCellStr(sheet, "F"+N, "Hello World")
		xlsx.SetCellStr(sheet, "G"+N, "Hello World")
		xlsx.SetCellStr(sheet, "H"+N, "Hello World")
		xlsx.SetCellStr(sheet, "I"+N, "Hello World")
		xlsx.SetCellStr(sheet, "J"+N, "Hello World")
		xlsx.SetCellStr(sheet, "K"+N, "Hello World")
		xlsx.SetCellStr(sheet, "L"+N, "Hello World")
		xlsx.SetCellStr(sheet, "M"+N, "Hello World")
		xlsx.SetCellStr(sheet, "N"+N, "Hello World")
		xlsx.SetCellStr(sheet, "O"+N, "Hello World")
	}
	xlsx.SaveAs("reports/xls1.xlsx")
}

func xls2Save(xlsx *xlsx2.File) {
	sheet := xlsx.Sheets[0]
	for i := 0; i < 15000; i++ {
		row := sheet.AddRow()
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
		row.AddCell().SetString("Hello World")
	}
	xlsx.Save("reports/xls2.xlsx")
}
