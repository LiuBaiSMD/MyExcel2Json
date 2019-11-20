package handler

import (
	"testing"
	"fmt"
	"github.com/tealeg/xlsx"
)

func TestGtOtherSheetInfo(t *testing.T){
	excelFileName := "../example/Sample.xlsx"
	var Xfile *xlsx.File
	var err error
	Xfile, err = xlsx.OpenFile(excelFileName)
	tt := "SheetInfo"
	result, err :=GetOtherSheetInfo(Xfile, tt, tt, "SheetName:Sample|KeyID:ID|FindID:102")
	if result == nil{
		fmt.Println("not find value")
	}
	if tt == "[]SheetInfo"{
		for i, j:= range result.([]interface{}){
			fmt.Println(i, " : ", j, err)
		}
	}
	if tt == "SheetInfo"{
		fmt.Println("map : ", result, err)
	}
}
