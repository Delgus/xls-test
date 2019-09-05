package main

import (
	"bytes"
	"fmt"
	"strconv"
	"sync"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	xlsx2 "github.com/tealeg/xlsx"
)

func BenchmarkFirstExcel(b *testing.B) {
	var wg sync.WaitGroup
	wg.Add(b.N)
	for i := 0; i < b.N; i++ {
		go func() {
			defer wg.Done()
			xlsx, err := excelize.OpenFile("resources/turnover.xlsx")
			if err != nil {
				fmt.Println("oops:", err)
				return
			}
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
			writer := new(bytes.Buffer)
			if err := xlsx.Write(writer); err != nil {
				fmt.Println("oops", err)
			}
		}()
	}
	wg.Wait()
}

func BenchmarkSecondExcel(b *testing.B) {
	var wg sync.WaitGroup
	wg.Add(b.N)
	for i := 0; i < b.N; i++ {
		go func() {
			defer wg.Done()
			xlsx, err := xlsx2.OpenFile("resources/turnover.xlsx")
			if err != nil {
				fmt.Println("oops:", err)
				return
			}
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
			writer := new(bytes.Buffer)
			if err := xlsx.Write(writer); err != nil {
				fmt.Println("oops", err)
			}
		}()
	}
	wg.Wait()
}
