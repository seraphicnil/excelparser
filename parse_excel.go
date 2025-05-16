package main

import (
	"encoding/json"
	"errors"
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
	Interrupt bool
)

func ParseOnFile(fileName string) error {
	if strings.Contains(fileName, ".xlsx") == false {
		return errors.New("invalid excel file")
	}
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		fmt.Println(err)
		return err
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
				err = parseSheet(f, excelSheetName)
				if err != nil {
					if Interrupt {
						return err
					}
					continue
				}
			}
		}
	}
	return nil
}

func parseSheet(f *excelize.File, sheetName string) error {
	// 定义字段名和类型
	fields := make([]reflect.StructField, 0)

	excelSheetName := "export_" + sheetName
	rows, err := f.GetRows(excelSheetName)
	if err != nil {
		fmt.Println(err)
		return err
	}
	row0 := rows[0]
	row1 := rows[1]
	row2 := rows[2]

	invalidIdxMap := make(map[int]bool)
	for idx, field := range row0 {
		exportType := row2[idx]
		if exportType == "c" || exportType == "C" {
			invalidIdxMap[idx] = true
			continue
		}
		fieldType := row1[idx]
		fields = append(fields, reflect.StructField{Name: field, Type: getType(fieldType)})
	}

	// 创建结构体类型
	structType := reflect.StructOf(fields)

	exportFileName := sheetName + ".export"
	if len(ExportDir) > 0 {
		os.RemoveAll(ExportDir)
		if _, err := os.Stat(ExportDir); os.IsNotExist(err) {
			os.MkdirAll(ExportDir, os.ModePerm)
		}
		exportFileName = ExportDir + "/" + exportFileName
	}
	_, err = os.Create(exportFileName)
	if err != nil {
		fmt.Println(err)
		return err
	}
	for rowIndex, row := range rows {
		if rowIndex > 3 {
			// 创建该类型的实例
			instance := reflect.New(structType).Elem()
			idx := -1
			for colIndex, cellValue := range row {
				if _, ok := invalidIdxMap[colIndex]; ok {
					continue
				}
				idx++
				fieldType := row1[colIndex]
				switch fieldType {
				case "uint32":
					if len(cellValue) == 0 {
						instance.Field(idx).SetUint(0)
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(idx).SetUint(uint64(num64))
				case "uint64":
					if len(cellValue) == 0 {
						instance.Field(idx).SetUint(uint64(0))
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(idx).SetUint(uint64(num64))
				case "int32":
					if len(cellValue) == 0 {
						instance.Field(idx).SetInt(0)
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(idx).SetInt(num64)
				case "int64":
					if len(cellValue) == 0 {
						instance.Field(idx).SetInt((0))
						continue
					}
					num64, err := strconv.ParseInt(cellValue, 10, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					} else {
						//fmt.Printf("转换结果: %d, 类型: %T\n", num64, num64) // 输出: -456, int64
					}
					instance.Field(idx).SetInt(num64)
				case "string":
					instance.Field(idx).SetString(cellValue)
				case "float":
					f64, err := strconv.ParseFloat(cellValue, 64) // 参数：字符串, 进制(10), 位数(64)
					if err != nil {
						fmt.Println(err)
					}
					instance.Field(idx).SetFloat(f64)
				case "bool":
					if len(cellValue) == 0 {
						instance.Field(idx).SetBool(false)
					} else {
						instance.Field(idx).SetBool(true)
					}
				}
			}

			// 打印结果
			//fmt.Printf("%+v\n", instance.Interface())
			val, err := json.Marshal(instance.Interface())
			if err != nil {
				fmt.Println(err)
				return err
			}
			writeFile(exportFileName, val)
			writeFile(exportFileName, []byte("\n"))
		}
	}
	return nil
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
