package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
)

func main() {
	input := bufio.NewScanner(os.Stdin)
	fmt.Println("请输入文件名: ")
	input.Scan()

	file, err := excelize.OpenFile("C:\\Users\\admin\\Desktop\\" + input.Text())
	if err != nil {
		fmt.Println(err.Error())
		return
	}

	// table 1
	needKey := []int{0, 1, 2, 3}
	sheet1 := GetData(file, "Sheet1", needKey)

	// table 2
	needKey = []int{1, 2, 3}
	sheet2 := GetData(file, "Sheet2", needKey)
}

// @title    GetData
// @description   获取指定列的数据
// @param     file        *excelize.File       excel库文件载体
// @param     sheetName   sting                excel下方标签名称
// @param     needKey     []int                需要指定的列的KEY
// @return    newData     map[string][]string  获取后的数据
func GetData(file *excelize.File, sheetName string, needKey []int) map[string][]string {
	data, err := file.GetCols(sheetName)
	if err != nil {
		fmt.Println(err.Error())
		return make(map[string][]string)
	}

	newData := make(map[string][]string)
	columnName := ""
	// 循环列
	for columnKey, columnArray := range data {
		// 循环行
		for rowKey, value := range columnArray {
			if rowKey == 0 {
				columnName = value
			}
			// 第一列 从第二行开始为实际订单号
			if rowKey > 0 && InArray(columnKey, needKey) {
				newData[columnName] = append(newData[columnName], value)
			}
		}
	}
	return newData
}

// @title    InArray
// @description   判断值是否在数组中
// @param     value       int                  需要寻找的值
// @param     array       []int                寻找的数组
// @return    bool                             结果
func InArray(value int, array []int) bool {
	for _, rangeValue := range array {
		if value == rangeValue {
			return true
		}
	}
	return false
}
