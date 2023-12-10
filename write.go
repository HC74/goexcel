package goexcel

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
)

const (
	DefaultSheetName = "Sheet1"
)

type WriteExcel[T any] struct {
	file      *excelize.File
	m         *MiddleData[T]
	num       int64       // 行数
	tw        *TypeWriter // 打字机
	sheetName string      // sheet的名字
	disStream bool
}

func (w *WriteExcel[T]) NumAdd() {
	w.num++
}

func (w *WriteExcel[T]) SheetName(sheetName string) *WriteExcel[T] {
	w.sheetName = sheetName
	return w
}

func (w *WriteExcel[T]) SetCellValue(name string, v any) {
	err := w.file.SetCellValue(w.sheetName, name, v)
	if err != nil {
		return
	}
}

func (w *WriteExcel[T]) NumS() string {
	return strconv.FormatInt(w.num, 10)
}

func (w *WriteExcel[T]) BuildForStream() *WriteExcel[T] {
	mid := w.m
	f := w.file

	sw, err := f.NewStreamWriter(w.sheetName)
	if err != nil {
		panic(err.Error())
	}

	var headers = make([]interface{}, 0, len(mid.Headers))

	// 创建header
	for _, header := range mid.Headers {
		columnName := findColumnName(header.field, header.TagName)
		headers = append(headers, excelize.Cell{Value: columnName})
	}
	err = sw.SetRow(w.tw.Next()+w.NumS(), headers)
	if err != nil {
		panic(err.Error())
	}

	w.tw.Clear()
	w.NumAdd()

	// 大规模数据填充
	var tempValue any

	var cellName string

	// 填充数据
	for {
		tempValue = mid.Values.Front()
		if tempValue == nil {
			break
		}
		// 填充
		results := tempValue.([]interface{})
		prefix := w.tw.Next()
		// 单元格名
		cellName = prefix + w.NumS()

		err := sw.SetRow(cellName, results)
		if err != nil {
			fmt.Println(err.Error())
		}
		// 行数迭代
		w.NumAdd()
		w.tw.Clear()
	}

	err = sw.Flush()
	if err != nil {
		panic(err.Error())
	}
	return w
}

func (w *WriteExcel[T]) DisStream() *WriteExcel[T] {
	w.disStream = true
	return w
}

func (w *WriteExcel[T]) Build() *WriteExcel[T] {
	mid := w.m
	f := w.file

	sheetIndex, err := f.NewSheet(w.sheetName)
	if err != nil {
		panic(err.Error())
		return w
	}

	f.SetActiveSheet(sheetIndex)

	// 自定义sheetName
	if !strings.EqualFold(DefaultSheetName, w.sheetName) {
		// 删除默认的
		err = f.DeleteSheet(DefaultSheetName)
		if err != nil {
			panic("删除sheet1失败")
		}
	}

	// 判断数据量，以决定是否需要流式处理
	if w.disStream == false && w.m.Values.size >= 100000 {
		// 数据量超过十万，启用流式处理
		return w.BuildForStream()
	}

	var cellName string

	// 创建header
	for _, header := range mid.Headers {
		columnName := findColumnName(header.field, header.TagName)
		prefix := w.tw.Next()
		// 单元格名
		cellName = prefix + w.NumS()
		w.SetCellValue(cellName, columnName)
	}

	// 重置打字机
	w.tw.Clear()

	// 行数迭代
	w.NumAdd()

	var tempValue any

	// 填充数据
	for {
		tempValue = mid.Values.Front()
		if tempValue == nil {
			break
		}
		// 填充
		results, err := convertToSliceOfStrings(tempValue)
		if err != nil {
			continue
		}
		for _, value := range results {
			prefix := w.tw.Next()
			// 单元格名
			cellName = prefix + w.NumS()
			w.SetCellValue(cellName, value)
		}
		// 行数迭代
		w.NumAdd()
		w.tw.Clear()
	}
	return w
}

func (w *WriteExcel[T]) SaveAs(name string) error {
	defer w.file.Close()
	err := w.file.SaveAs(name)
	if err != nil {
		return err
	}

	return nil
}
