package goexcel

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strconv"
)

type ReadExcel[T any] struct {
	file      *excelize.File
	sheetName string
	v         T
	mTitle    map[int]any
}

// SheetName 设置sheet名
func (r *ReadExcel[T]) SheetName(sheetName string) *ReadExcel[T] {
	r.sheetName = sheetName
	return r
}

func (r *ReadExcel[T]) Read(v any) error {
	destType := reflect.TypeOf(v)
	if destType.Kind() != reflect.Ptr || destType.Elem().Kind() != reflect.Slice {
		return fmt.Errorf("v must be a pointer to a slice")
	}
	rowsa, _ := r.file.GetRows(r.sheetName)
	// 减去title
	rowLen := len(rowsa) - 1
	rowsa = nil
	// 获取切片元素类型
	sliceElemType := destType.Elem().Elem()

	// 创建结构体切片实例
	slice := reflect.MakeSlice(destType.Elem(), rowLen, rowLen)

	rows, err := r.file.Rows(r.sheetName)
	if err != nil {
		panic(err.Error())
	}
	err = r.Title(rows)
	if err != nil {
		return err
	}
	idx := 0
	for rows.Next() {
		columns, err := rows.Columns()
		if err != nil {
			return err
		}
		structInstance := reflect.New(sliceElemType).Elem()
		tagMap := make(map[string]string)
		for i, column := range columns {
			value, isok := r.mTitle[i]
			if isok {
				tagName := value.(string)
				tagMap[tagName] = column
			}
		}
		// 遍历结构体字段并设置值
		for j := 0; j < sliceElemType.NumField(); j++ {
			field := sliceElemType.Field(j)
			tagName := field.Tag.Get("ex") // 替换为你的结构体标签

			// 如果标签匹配，则设置字段值
			if tagName == "-" {
				continue
			}

			tagName = getExTag(tagName, "name")
			value, isok := tagMap[tagName]
			if isok {
				switch field.Type.Kind() {
				case reflect.String:
					structInstance.FieldByName(field.Name).SetString(value)
				case reflect.Int64, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32:
					v, _ := strconv.ParseInt(value, 10, 64)
					structInstance.FieldByName(field.Name).SetInt(v)
				}

			}
		}
		// 将结构体实例放入切片
		slice.Index(idx).Set(structInstance)
		idx++
	}
	// 将填充好的切片设置回传入的指针
	reflect.ValueOf(v).Elem().Set(slice)

	return nil
}

func (r *ReadExcel[T]) Title(rows *excelize.Rows) error {
	// 拿到title
	exist := rows.Next()

	if !exist {
		return errors.New("无数据")
	}

	columns, err := rows.Columns()
	if err != nil {
		return err
	}

	var mTitle = make(map[int]any)
	for i, title := range columns {
		mTitle[i] = title
	}
	r.mTitle = mTitle
	return nil
}
