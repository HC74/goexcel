package goexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
	"testing"
	"time"
)

type TestData struct {
	A string `ex:"name:A列"`
	B string `ex:"name:B列"`
	C string `ex:"name:C列"`
	D string `ex:"name:D列"`
	E string `ex:"name:E列"`
	F string `ex:"name:F列"`
	G string `ex:"name:G列"`
	H string `ex:"name:H列"`
	L string `ex:"-"`
	N int64  `ex:"name:N列"`
}

func TestC(t *testing.T) {
	var results []TestData
	err := Load("test.xlsx", TestData{}).SheetName("Sheet2").Read(&results)
	if err != nil {
		return
	}
	fmt.Println("end,")
}

func TestB(t *testing.T) {
	num := 10
	var datas = make([]TestData, 0, num)
	for i := 0; i < num; i++ {
		is := strconv.Itoa(i + 1)
		datas = append(datas, TestData{
			A: "A" + is,
			B: "B" + is,
			C: "C" + is,
			D: "D" + is,
			E: "E" + is,
			F: "F" + is,
			G: "G" + is,
			H: "H" + is,
			L: "L" + is,
			N: int64(i),
		})
	}
	startTime := time.Now()
	err := New(datas, TestData{}).SheetName("Sheet2").Build().
		SaveAs("test.xlsx")
	elapsedTime := time.Since(startTime)
	fmt.Printf("Time taken: %s\n", elapsedTime)
	if err != nil {
		fmt.Println(err.Error())
		fmt.Println("??")
	}
	//fmt.Println("++++")
	//GetSlice(datas2, TestData{})
}

// Time taken: 43.012463333s
// Time taken: 43.239224458s
func TestAA(t *testing.T) {
	num := 1000000
	var datas = make([]TestData, 0, num)
	for i := 0; i < num; i++ {
		is := strconv.Itoa(i + 1)
		datas = append(datas, TestData{
			A: "A" + is,
			B: "B" + is,
			C: "C" + is,
			D: "D" + is,
			E: "E" + is,
			F: "F" + is,
			G: "G" + is,
			H: "H" + is,
			L: "L" + is,
		})
	}
	startTime := time.Now()
	f := excelize.NewFile()
	sheet, _ := f.NewSheet("Sheet2")
	f.SetActiveSheet(sheet)
	_ = f.DeleteSheet("Sheet1")
	headers := []string{"A列", "B列", "C列", "D列", "E列", "F列", "G列", "H列"}
	c := 'A'
	for _, header := range headers {
		_ = f.SetCellValue("Sheet2", string(c)+strconv.Itoa(1), header)
		c = c + 1
	}
	for i, data := range datas {
		_ = f.SetCellValue("Sheet2", "A"+strconv.Itoa(i+2), data.A)
		_ = f.SetCellValue("Sheet2", "B"+strconv.Itoa(i+2), data.B)
		_ = f.SetCellValue("Sheet2", "C"+strconv.Itoa(i+2), data.C)
		_ = f.SetCellValue("Sheet2", "D"+strconv.Itoa(i+2), data.D)
		_ = f.SetCellValue("Sheet2", "E"+strconv.Itoa(i+2), data.E)
		_ = f.SetCellValue("Sheet2", "F"+strconv.Itoa(i+2), data.F)
		_ = f.SetCellValue("Sheet2", "G"+strconv.Itoa(i+2), data.G)
		_ = f.SetCellValue("Sheet2", "H"+strconv.Itoa(i+2), data.H)
	}
	f.SaveAs("test2.xlsx")
	elapsedTime := time.Since(startTime)
	fmt.Printf("Time taken: %s\n", elapsedTime)
}

func TestA(t *testing.T) {
	writer := TypeWriter{}
	for i := 1; i <= 5; i++ {
		for j := 1; j <= 5; j++ {
			fmt.Print(writer.Next() + strconv.Itoa(i) + ",")
		}
		fmt.Println()
		writer.Clear()
	}
}
