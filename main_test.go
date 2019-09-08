package main

import (
	"bytes"
	"fmt"
	"strconv"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	xlsx2 "github.com/tealeg/xlsx"
)

//library excelize low
func BenchmarkExcelizeLibLow(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := excelize.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
		for y := 0; y < 15; y++ {
			N := strconv.Itoa(y + 1)
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
	}
}

//library xlsx low
func BenchmarkXLSXLibLow(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := xlsx2.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.Sheets[0]
		for y := 0; y < 15; y++ {
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
	}
}

//library excelize high
func BenchmarkExcelizeLibHigh(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := excelize.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
		for y := 0; y < 15000; y++ {
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
	}
}

//library xlsx high
func BenchmarkXLSXLibHigh(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := xlsx2.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.Sheets[0]
		for y := 0; y < 15000; y++ {
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
	}
}

//library excelize high parallel
func BenchmarkExcelizeLibHighParallel(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		xlsx, err := excelize.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}

		sheet := xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
		for y := 0; y < 15000; y++ {
			N := strconv.Itoa(y + 1)
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
		pb.Next()
	})
}

//library xlsx high parallel
func BenchmarkXLSXLibHighParallel(b *testing.B) {
	b.RunParallel(func(pb *testing.PB) {
		xlsx, err := xlsx2.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}

		sheet := xlsx.Sheets[0]
		for y := 0; y < 15000; y++ {
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
		writer := new(bytes.Buffer)
		if err := xlsx.Write(writer); err != nil {
			fmt.Println("oops", err)
		}
		pb.Next()
	})
}

//library excelize file
func BenchmarkExcelizeLibFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := excelize.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.GetSheetName(xlsx.GetActiveSheetIndex())
		for y := 0; y < 15000; y++ {
			N := strconv.Itoa(y + 1)
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
		nameNumber := strconv.Itoa(i + 1)
		if err := xlsx.SaveAs("bench-excelize/xls" + nameNumber + ".xlsx"); err != nil {
			fmt.Println("oops", err)
		}
	}
}

//library xlsx file
func BenchmarkXLSXLibFile(b *testing.B) {
	for i := 0; i < b.N; i++ {
		xlsx, err := xlsx2.OpenFile("resources/turnover.xlsx")
		if err != nil {
			fmt.Println("oops:", err)
			return
		}
		sheet := xlsx.Sheets[0]
		for y := 0; y < 15000; y++ {
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
		nameNumber := strconv.Itoa(i + 1)
		if err := xlsx.Save("bench-xlsx/xls" + nameNumber + ".xlsx"); err != nil {
			fmt.Println("oops", err)
		}
	}
}
