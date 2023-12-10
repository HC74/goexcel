package goexcel

import (
	"github.com/xuri/excelize/v2"
	"reflect"
	"strings"
)

func New[T any](data []T, t T) *WriteExcel[T] {
	file := excelize.NewFile()
	m := GetReflectSlice(data, t)
	w := &WriteExcel[T]{
		file:      file,
		m:         m,
		num:       1,
		tw:        &TypeWriter{},
		sheetName: "Sheet1",
		disStream: false,
	}
	return w
}

// Load 加载文件
func Load[T any](fileName string, t T) *ReadExcel[T] {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		panic(err.Error())
	}
	return &ReadExcel[T]{
		file:      f,
		sheetName: DefaultSheetName,
		v:         t,
	}
}

func findColumnName(v reflect.StructField, tag string) string {
	columnName := getExTag(tag, "name")
	if strings.EqualFold(columnName, "") {
		// 未找到，使用属性名
		return v.Name
	}
	return columnName
}
