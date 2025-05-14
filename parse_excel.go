package main

import (
	"encoding/json"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"reflect"
	"strconv"
	"strings"
)

var (
	SourceDir string
	ExportDir string
)

func ParseOnFile(fileName string) {
	if strings.Contains(fileName, ".xlsx") == false {
		return
	}
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return
	}
	defer f.Close() // 确保关闭文件句柄^[3]
	fmt.Printf("excel:  %s\n", fileName)
	excelSheetName := ""

	sheets := f.GetSheetMap()
	for idx, name := range sheets {
		fmt.Printf("工作表 %d: %s\n", idx, name)
		nameArr := strings.Split(name, "_")
		if len(nameArr) == 2 {
			if nameArr[0] == "export" {
				excelSheetName = nameArr[1]
				parseSheet(f, excelSheetName)
			}
		}
	}
}

func parseSheet(f *excelize.File, sheetName string) {
	// 定义字段名和类型
	fields := make([]reflect.StructField, 0)

	excelSheetName := "export_" + sheetName
	rows, err := f.GetRows(excelSheetName)
	if err != nil {
		fmt.Println(err)
		return
	}
	row0 := rows[0]
	row1 := rows[1]

	for idx, field := range row0 {
		fieldType := row1[idx]
		fields = append(fields, reflect.StructField{Name: field, Type: getType(fieldType)})
	}

	// 创建结构体类型
	structType := reflect.StructOf(fields)

	exportFileName := sheetName + ".export"
	if len(ExportDir) > 0 {
		if _, err := os.Stat(exportFileName); os.IsNotExist(err) {
			os.MkdirAll(ExportDir, os.ModePerm)
		}
		exportFileName = ExportDir + "/" + exportFileName
	}
	os.Create(exportFileName)
	for rowIndex, row := range rows {
		if rowIndex > 2 {
			// 创建该类型的实例
			instance := reflect.New(structType).Elem()
			for colIndex, cellValue := range row {
				fieldType := row1[colIndex]
				switch fieldType {
				case "uint32":
					if len(cellValue) == 0 {
						instance.Field(colIndex).SetUint(0)
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(colIndex).SetUint(uint64(num64))
				case "uint64":
					if len(cellValue) == 0 {
						instance.Field(colIndex).SetUint(uint64(0))
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(colIndex).SetUint(uint64(num64))
				case "int32":
					if len(cellValue) == 0 {
						instance.Field(colIndex).SetInt(0)
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(colIndex).SetInt(num64)
				case "int64":
					if len(cellValue) == 0 {
						instance.Field(colIndex).SetInt((0))
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(colIndex).SetInt(num64)
				case "string":
					instance.Field(colIndex).SetString(cellValue)
				case "float":
					f64, err := strconv.ParseFloat(cellValue, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					}
					instance.Field(colIndex).SetFloat(f64)
				case "bool":
					if len(cellValue) == 0 {
						instance.Field(colIndex).SetBool(false)
					} else {
						instance.Field(colIndex).SetBool(true)
					}
				}
			}

			// 打印结果
			//fmt.Printf("%+v\n", instance.Interface())
			val, err := json.Marshal(instance.Interface())
			if err != nil {
				fmt.Println(err)
			}
			writeFile(exportFileName, val)
			writeFile(exportFileName, []byte("\n"))
		}
	}
}

func getType(val string) reflect.Type {
	switch val {
	case "uint32":
		return reflect.TypeOf(uint32(0))
	case "uint64":
		return reflect.TypeOf(uint64(0))
	case "int32":
		return reflect.TypeOf(int32(0))
	case "int64":
		return reflect.TypeOf(int64(0))
	case "string":
		return reflect.TypeOf("")
	case "float":
		return reflect.TypeOf(float64(0))
	case "bool":
		return reflect.TypeOf(false)
	}
	return reflect.TypeOf("")
}

func writeFile(fileName string, data []byte) {
	f, err := os.OpenFile(fileName, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		log.Fatal(err)
	}
	if _, err := f.Write(data); err != nil {
		f.Close() // ignore error; Write error takes precedence
		log.Fatal(err)
	}
	if err := f.Close(); err != nil {
		log.Fatal(err)
	}
}

func readFile(fileName string) string {
	val, err := os.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	return string(val)
}
