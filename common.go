package goexcel

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"
)

type (
	// StructInfo 结构体详情
	StructInfo struct {
		field   reflect.StructField
		value   any
		TagName string
	}

	// TypeWriter 打字机
	TypeWriter struct {
		value  byte
		prefix string
		onOof  bool
	}

	// Node 节点
	Node[T any] struct {
		Next  *Node[T]
		value T
	}

	// LinkedList 单链表
	LinkedList[T any] struct {
		head *Node[T]
		tail *Node[T]
		size int64
	}

	MiddleData[T any] struct {
		Headers []StructInfo
		Values  *LinkedList[[]any]
	}
)

func NewLinkedList[T any]() *LinkedList[T] {
	return &LinkedList[T]{
		size: 0,
	}
}

func (ll *LinkedList[T]) Push(data T) {
	ll.size++
	newNode := &Node[T]{value: data, Next: nil}
	if ll.head == nil {
		ll.head = newNode
		ll.tail = newNode
		return
	}
	ll.tail.Next = newNode
	ll.tail = newNode
}

// Front pops the first element from the linked list
func (ll *LinkedList[T]) Front() interface{} {
	if ll.head == nil {
		return nil
	}
	ll.size--
	data := ll.head.value
	ll.head = ll.head.Next
	if ll.head == nil {
		ll.tail = nil // 如果链表为空，将尾结点设置为 nil
	}
	return data
}

func buildStructInfo(f reflect.StructField, v any) StructInfo {
	return StructInfo{
		field: f,
		value: v,
	}
}

// ReflectGetStructInfos 反射获取结构体信息
func ReflectGetStructInfos[T any](t T) []StructInfo {
	tValue := reflect.ValueOf(t)
	var results []StructInfo
	for i := 0; i < tValue.NumField(); i++ {
		field := tValue.Type().Field(i)
		fieldValue := tValue.Field(i).Interface()
		temp := buildStructInfo(field, fieldValue)
		temp.TagName = field.Tag.Get("ex")
		if temp.TagName == "-" {
			continue
		}
		results = append(results, temp)
	}
	return results
}
func GetReflectSlice[T any](data []T, t T) *MiddleData[T] {
	sliceValue := reflect.ValueOf(data)

	if sliceValue.Kind() != reflect.Slice {
		fmt.Println("Input is not a slice.")
		return nil
	}

	header := ReflectGetStructInfos(t)
	result := &MiddleData[T]{
		Headers: header,
	}

	if sliceValue.Len() <= 0 {
		return result
	}

	// 批量获取所有元素
	elements := make([]reflect.Value, sliceValue.Len())
	for i := 0; i < sliceValue.Len(); i++ {
		elements[i] = sliceValue.Index(i)
	}

	node := NewLinkedList[[]any]()

	startTime := time.Now()
	fmt.Println("action")
	// 遍历元素并构建结果
	for _, element := range elements {
		// 在每次循环中创建新的切片
		temps := make([]any, 0, len(header))
		fieldSize := element.NumField()
		for j := 0; j < fieldSize; j++ {
			field := element.Type().Field(j)
			tagName := field.Tag.Get("ex")
			if tagName == "-" {
				continue
			}
			fieldValue := element.Field(j).Interface()
			temps = append(temps, fieldValue)
		}
		node.Push(temps)
	}

	elapsedTime := time.Since(startTime)
	fmt.Printf("Time taken: %s\n", elapsedTime)

	result.Values = node
	return result
}

// Clear 缓冲
func (t *TypeWriter) Clear() {
	t.onOof = false
	t.prefix = ""
	t.value = 0
}

// Next 下一个字母
func (t *TypeWriter) Next() string {
	if t.value == 0 {
		t.value = byte('A')
		return string(t.value)
	}
	if t.onOof {
		res, d := t.execOverflow()
		t.value = d.value
		t.prefix = d.prefix
		t.value = t.value + 1
		return res
	}
	if t.value == byte('Z') {
		// 进入二级迭代
		t.onOof = true
		res, d := t.execOverflow()
		t.value = d.value
		t.prefix = d.prefix
		return res
	}
	t.value = t.value + 1
	return string(t.value)
}

// execOverflow 处理移除问题
func (t *TypeWriter) execOverflow() (string, *TypeWriter) {
	if t.prefix == "" {
		// 第一次晋级
		t.prefix = "A"
		t.value = 'A'
		return t.prefix + string(t.value), t
	}
	if t.value == '[' {
		nextPrefixValue := t.prefix[0] + 1
		t.value = 'A'
		t.prefix = string(nextPrefixValue)
	}
	return t.prefix + string(t.value), t
}

// getExTag 获取easyExcel对应的tag
func getExTag(tagName, flag string) string {
	if !strings.Contains(tagName, ";") {
		// 非多个tag
		return getSplitTag(tagName, flag)
	}

	// 有多个tag
	tags := strings.Split(tagName, ";")
	// 临时tag变量
	var tempTag string

	for _, tag := range tags {
		tempTag = getSplitTag(tag, flag)
		if !strings.EqualFold(tempTag, "") {
			return tempTag
		}
	}
	// 未找到
	return ""
}

func getSplitTag(v, key string) string {
	if !strings.Contains(v, ":") {
		// 不包含自定义，直接返回
		return ""
	}

	// 根据: 分割
	tagSplits := strings.Split(v, ":")
	// 根据Key查找对应value
	if tagSplits[0] == key {
		return tagSplits[1]
	}
	// 未找到对应value 返回空
	return ""
}

func convertToSliceOfStrings(input interface{}) ([]string, error) {
	// 将外层的 interface{} 转换为 []interface{}
	sliceInterface, ok := input.([]interface{})
	if !ok {
		return nil, fmt.Errorf("Input is not []interface{}")
	}

	// 遍历 []interface{}，将每个元素转换为 string
	result := make([]string, len(sliceInterface))
	for i, item := range sliceInterface {
		// 将每个元素转换为 string
		str, err := convertToString(item)
		if err != nil {
			return nil, err
		}
		result[i] = str
	}

	return result, nil
}

func convertToString(value interface{}) (string, error) {
	switch v := value.(type) {
	case int:
		return strconv.Itoa(v), nil
	case int64:
		return strconv.FormatInt(v, 10), nil
	case float64:
		return strconv.FormatFloat(v, 'f', -1, 64), nil
	case string:
		return v, nil
	default:
		return "", fmt.Errorf("Unsupported type: %T", v)
	}
}
