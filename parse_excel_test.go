package main

import (
	"encoding/json"
	"fmt"
	"strings"
	"testing"
)

func TestExample(t *testing.T) {
	fileName := "source/example.xlsx"
	ParseOnFile(fileName)
	type ExcelExample struct {
		Id            uint64
		ShopId        uint32
		Resource      string
		ExchangeMoney uint32
		ActId         uint32
		Lang          string
		Page          int32
		PageOpen      string
		Msg           string
		GoodType      int32
		ClientContent string
		GoodContent   string
	}

	exportFileName := "shop.export"
	content := readFile(exportFileName)
	contentList := strings.Split(content, "\n")
	for _, value := range contentList {
		if len(value) == 0 {
			continue
		}
		e := &ExcelExample{}
		err := json.Unmarshal([]byte(value), e)
		if err != nil {
			fmt.Println(err)
		}
		fmt.Printf("ExcelExample: %+v\n", e)
	}
}
