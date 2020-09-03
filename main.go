package main

import (
	"bufio"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"os"
	"strconv"
	"strings"
	"time"
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
	needName := []string{"订单编号", "买家实际支付金额", "买家留言", "收货人姓名", "收货地址", "联系手机", "宝贝标题", "订单备注"}
	sheet1 := GetDataByColumn(file, "Sheet1", needName)

	// table 2
	needName = []string{"订单编号", "商品属性"}
	sheet2 := GetDataByRow(file, "Sheet2", needName)

	// 创建新的表
	newFile := excelize.NewFile()

	// 循环数据写入新的表格
	var axis int
	var option string
	for columnKey, columnValue := range sheet1 {
		// 根据当前循环的key 计算字母坐标
		axis = 65 + columnKey
		// 循环列值
		for rowKey, rowValue := range columnValue {
			err = newFile.SetCellValue("Sheet1", string(axis)+strconv.Itoa(rowKey+1), rowValue)
			if err != nil {
				fmt.Println(err.Error())
			}
			// 如果是订单编号 遍历sheet2 获取对应数据
			if columnValue[0] == "订单编号" {
				if rowKey == 0 {
					err = newFile.SetCellValue("Sheet1", string(65+len(sheet1))+strconv.Itoa(rowKey+1), "商品属性")
					if err != nil {
						fmt.Println(err.Error())
					}
					continue
				}
				option = ""
				for _, Value := range sheet2 {
					if rowValue == Value[0] {
						option = option + Value[1] + " "
					}
				}
				err = newFile.SetCellValue("Sheet1", string(65+len(sheet1))+strconv.Itoa(rowKey+1), option)
				if err != nil {
					fmt.Println(err.Error())
				}
			}
		}
	}
	// 保存新文件
	err = newFile.SaveAs("C:\\Users\\admin\\Desktop\\" + time.Now().Format("2006-01-02") + ".xlsx")
	if err != nil {
		fmt.Println(err.Error())
	}
}

// @title    GetDataByColumn
// @description   获取指定列的数据
// @param     file        *excelize.File       excel库文件载体
// @param     sheetName   sting                excel下方标签名称
// @param     needName     []string            需要指定的列的名称
// @return    newData     map[string][]string  获取后的数据
func GetDataByColumn(file *excelize.File, sheetName string, needName []string) [][]string {
	data, err := file.GetCols(sheetName)
	if err != nil {
		fmt.Println(err.Error())
		return [][]string{}
	}

	var newData [][]string
	// 循环列
	for _, columnValue := range data {
		// 判断列名是否需要
		if !InArrayString(columnValue[0], needName) {
			continue
		}
		newData = append(newData, columnValue)
	}
	return newData
}

// @title    GetDataByRow
// @description   获取指定行的数据
// @param     file        *excelize.File       excel库文件载体
// @param     sheetName   sting                excel下方标签名称
// @param     needName     []string            需要指定的列的名称
// @return    newData     map[string][]string  获取后的数据
func GetDataByRow(file *excelize.File, sheetName string, needName []string) [][]string {
	data, err := file.GetRows(sheetName)
	if err != nil {
		fmt.Println(err.Error())
		return [][]string{}
	}

	var newData [][]string
	var keys []int
	var values []string
	// 循环行
	for rowKey, rowValue := range data {
		if rowKey == 0 {
			for key, value := range rowValue {
				if InArrayString(value, needName) {
					keys = append(keys, key)
				}
			}
		}
		values = []string{}
		for key, value := range rowValue {
			if InArrayInt(key, keys) {
				values = append(values, value)
			}
		}
		newData = append(newData, values)
	}
	return newData
}

// @title    InArray
// @description   判断值是否在数组中
// @param     value       string               需要寻找的名称
// @param     array       []string             寻找的数组
// @return    bool                             结果
func InArrayString(value string, array []string) bool {
	for _, rangeValue := range array {
		if strings.TrimSpace(value) == rangeValue {
			return true
		}
	}
	return false
}

// @title    InArray
// @description   判断值是否在数组中
// @param     value       int               需要寻找的名称
// @param     array       []int             寻找的数组
// @return    bool                             结果
func InArrayInt(value int, array []int) bool {
	for _, rangeValue := range array {
		if value == rangeValue {
			return true
		}
	}
	return false
}
