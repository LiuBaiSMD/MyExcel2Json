package handler

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"strings"
	"strconv"
	"encoding/json"
	"errors"
	"os"
	"github.com/micro/go-micro/util/log"
)


func ExcelChanger(filePath string) {
	ifExist, _ :=  CheckPathExists(filePath)
	if !ifExist{
		panic("配表文件不存在请检查 ------ > "+ filePath)
		return
	}
	excelFileName := filePath
	var xlFile *xlsx.File
	var err error
	xlFile, err = xlsx.OpenFile(excelFileName)
	if err != nil {
		panic(err)
	}

	for _, sheet := range xlFile.Sheets {
		//第一个先把所有数据的Key得到
		if strings.HasPrefix(sheet.Name, "@"){
			fmt.Println("此表为子表不做处理------> ",sheet.Name)
			continue
		}
		var sheetResult []interface{}
		sheetResult = make([]interface{}, 0)
		for irow, row := range sheet.Rows {
			ifAdd := true
			var rowRessult map[string]interface{}
			rowRessult = make(map[string]interface{}, 0)
			//ID := row.Cells[0].String()
			for col, cell := range row.Cells {
				if irow <= 3 {
					ifAdd = false
					continue
				}
				//开始构造结果体， 判断是否数据缺失
				if len(sheet.Rows[0].Cells) < col+ 1{
					panic("表格式错误")
				}
				key := sheet.Rows[0].Cells[col].String()
				valType :=  sheet.Rows[1].Cells[col].String()
				desc := sheet.Rows[2].Cells[col].String()
				var value interface{}
				//进行类型判断,并转换
				value ,_ = changeString2Type(xlFile, cell.String(), valType, desc)
				if key != ""{
					rowRessult[key] = value
				}
			}
			if ifAdd{//判断是否加入result集合
				sheetResult = append(sheetResult, rowRessult)
				var test sheetInfo
				jsonStr, _ := json.Marshal(rowRessult)
				json.Unmarshal([]byte(jsonStr), &test)
				fmt.Println("scan: ", test)
			}
		}
		err = StoreWithJson(sheet.Name, sheetResult)
		if err != nil{
			panic(err)
		}
		log.Log("sheet表 -----> ["+ sheet.Name + "] 处理完毕！"  )
	}
}

func getIntList(decs, value string)([]int, error){
	Spliter := ","
	strList := strings.Split(value, Spliter)
	var intList []int
	intList = make([]int, 0)
	for _, j:= range strList{
		if j==""{
			continue
		}
		intStr, err := strconv.Atoi(j)
		if err != nil{
			panic(value + "出现错误" + err.Error())
			return nil, err
		}
		intList = append(intList,intStr)
	}
	return intList, nil
}

func getInt64List(decs, value string)([]int64, error){
	Spliter := ","
	strList := strings.Split(value, Spliter)
	var intList []int64
	intList = make([]int64, 0)
	for _, j:= range strList{
		if j==""{
			continue
		}
		intStr, err := strconv.ParseInt(j, 10, 64)
		if err != nil{
			panic(value + "出现错误" + err.Error())
			return nil, err
		}
		intList = append(intList,intStr)
	}
	return intList, nil
}

func getInt(decs, value string)(int, error){
	dft := 0
	if value==""{
		return dft, nil
	}
	i, err := strconv.Atoi(value)
	if err != nil{
		panic(value + "出现错误" + err.Error())
		return dft, err
	}
	return i, nil
}

func getInt64(decs, value string)(int64, error){
	dft := (int64)(0)
	if value==""{
		return dft, nil
	}
	i, err := strconv.ParseInt(value, 10, 64)
	if err != nil{
		panic(value + "出现错误" + err.Error())
		return dft, err
	}
	return i, nil
}

func Str2Map(decs, value string)(map[string]interface{}, error){
	var mapResult map[string]interface{}
	err:= json.Unmarshal([]byte(value), &mapResult)
	return mapResult,err
}

func getStrList(desc, value string)([]string, error){
	spliter := "|"
	StrList := strings.Split(value, spliter)
	return StrList, nil
}

func GetOtherSheetInfo(Xfile *xlsx.File, valType, decs, cellStrValue string)(interface{}, error){
	//keyID为唯一识别的ID标识，如"ID" 或者"id"，等索引的标识服务，excel第一行中的数据
	infosMap := splitSheetInfo(cellStrValue)
	keyID := infosMap["KeyID"]
	SheetName := infosMap["SheetName"]
	findID := infosMap["FindID"]
	var result interface{}
	var err error
	if valType == "[]SheetInfo"{
		result, err = getOtherSheetInfoList(Xfile, keyID, decs, SheetName, findID )
	}
	if valType == "SheetInfo"{
		result, err = getOtherSheetInfoMap(Xfile, keyID, decs, SheetName, findID )
	}
	return result, err
}

func getOtherSheetInfoList(Xfile *xlsx.File, keyID, decs, sheetName, findID string)([]interface{}, error){
	//keyID为唯一识别的ID标识，如"ID" 或者"id"，等索引的标识服务，excel第一行中的数据
	sheet := *Xfile.Sheet[sheetName]
	//找出所有跟findID一样的ID
	lRows := len(sheet.Rows)
	if lRows<5{
		return nil, errors.New("excel表错误，缺少数据内容等！")
	}
	//找出存储ID的那一列
	lCols := len(sheet.Rows[0].Cells)
	//找到ID于findId一样的并记录
	keyIDCol := -1
	for col:=0;col<lCols;col++{
		if sheet.Rows[0].Cells[col].String() == keyID{
			keyIDCol = col
			break
		}
	}
	if keyIDCol==-1{
		return nil, errors.New(`查无 [`+ keyID+`] 列！`)
	}
	//找出所有ID= findID的值得行
	findIdRows := make([]int, 0)
	for i:=4;i< lRows;i++{
		if sheet.Rows[i].Cells[keyIDCol].String() == findID{
			findIdRows = append(findIdRows, i)
		}
	}
	var findIDResult []interface{}
	if len(findIdRows) == 0{
		return nil,errors.New(`查无 [`+ findID+`] 数据！`)
	}else{
		for _, IDRow := range findIdRows {
			//开始组装数据
			var rowRessult map[string]interface{}
			rowRessult = make(map[string]interface{}, 0)
			//ID := row.Cells[0].String()
			row := sheet.Rows[IDRow]
			for col, cell := range row.Cells {
				//开始构造结果体
				if len(sheet.Rows[0].Cells) < col+1 {
					panic("表格式错误")
				}
				key := sheet.Rows[0].Cells[col].String()
				valType := sheet.Rows[1].Cells[col].String()
				desc := sheet.Rows[2].Cells[col].String()
				var value interface{}
				//进行类型判断
				value, _ = changeString2Type(Xfile, cell.String(), valType, desc)
				if key != ""{
					rowRessult[key] = value
				}
			}
			findIDResult = append(findIDResult, rowRessult)
		}
	}
	return findIDResult, nil
}

func getOtherSheetInfoMap(Xfile *xlsx.File, keyID, decs, sheetName, findID string)(map[string]interface{}, error){
	//keyID为唯一识别的ID标识，如"ID" 或者"id"，等索引的标识服务，excel第一行中的数据
	sheet := *Xfile.Sheet[sheetName]
	//找出所有跟findID一样的ID
	lRows := len(sheet.Rows)
	if lRows<5{
		return nil, errors.New("excel表错误，缺少数据内容等！")
	}
	//找出存储ID的那一列
	lCols := len(sheet.Rows[0].Cells)
	//找到ID于findId一样的并记录
	keyIDCol := -1
	for col:=0;col<lCols;col++{
		if sheet.Rows[0].Cells[col].String() == keyID{
			keyIDCol = col
			break
		}
	}
	if keyIDCol == -1{
		return nil, errors.New(`查无 [`+ keyID+`] 列！`)
	}
	//找出所有ID= findID的值得行
	findIDRows := make([]int, 0)
	for i:=4;i< lRows;i++{
		if sheet.Rows[i].Cells[keyIDCol].String() == findID{
			findIDRows = append(findIDRows, i)
		}
	}
	var findIDResult map[string]interface{}
	if len(findIDRows) == 0 {
		return nil, errors.New(`查无 [`+ findID+`] 数据！`)
	}else if len(findIDRows) != 1{
		return nil, errors.New(`[`+ findID+`] 数据存在多条，请检查配置！`)
	} else{
		IDRow := findIDRows[0]
		//开始组装数据
		var rowRessult map[string]interface{}
		rowRessult = make(map[string]interface{}, 0)
		//ID := row.Cells[0].String()
		row := sheet.Rows[IDRow]
		for col, cell := range row.Cells {
			//开始构造结果体
			if len(sheet.Rows[0].Cells) < col+1 {
				panic("表格式错误")
			}
			key := sheet.Rows[0].Cells[col].String()
			valType := sheet.Rows[1].Cells[col].String()
			desc := sheet.Rows[2].Cells[col].String()
			var value interface{}
			//进行类型判断
			value, _ = changeString2Type(Xfile, cell.String(), valType, desc)
			if key != ""{
				rowRessult[key] = value
			}
		}
		findIDResult = rowRessult
	}
	return findIDResult, nil
}

func changeString2Type(Xfile *xlsx.File, cellStrValue, valtype , desc string)(interface{}, error){
	//进行类型判断
	var value interface{}
	var err = errors.New("")
	err = nil
	cellStrValue = strings.Replace(cellStrValue, " ", "", -1)
	if checkIfNilStr(cellStrValue){
		value = getDefaultVal(valtype)
		return value, nil
	}
	valtype = strings.Replace(valtype, " ", "", -1)
	switch valtype {
	case "[]int32":
		value, err = getIntList(desc, cellStrValue)
	case "[]int64":
		value, err = getInt64List(desc, cellStrValue)
	case "int":
		value, err = getInt(desc, cellStrValue)
	case "int64":
		value, err = getInt64(desc, cellStrValue)
	case "string":
		value = cellStrValue
	case "mapInterface":
		value, err = Str2Map(desc, cellStrValue)
	case "[]string":
		value, err = getStrList(desc, cellStrValue)
	case "[]SheetInfo":
		value, err = GetOtherSheetInfo(Xfile, "[]SheetInfo", desc, cellStrValue)
	case "SheetInfo":
		value, err = GetOtherSheetInfo(Xfile, "SheetInfo", desc, cellStrValue)
	default:
		value = cellStrValue
	}
	return value, err
}

func getDefaultVal(valtype string)interface{}{
	//对于空值设置初始值
	valtype = strings.Replace(valtype, " ", "", -1)
	var value interface{}
	switch valtype {
	case "[]int32":
		value = 0
	case "[]int64":
		value = []int64{}
	case "int":
		value = 0
	case "int64":
		value = (int64)(0)
	case "string":
		value = ""
	case "mapInterface":
		value = map[string]interface{}{}
	case "[]string":
		value = []string{}
	case "[]SheetInfo":
		value = []interface{}{}
	case "SheetInfo":
		value = map[string]interface{}{}
	default:
		value = ""
	}
	return value
}

func splitSheetInfo(sheetValue string)(map[string]string){
	sheetValue = strings.Replace(sheetValue, " ", "", -1)
	s := strings.Split(sheetValue, "|")
	if len(s)!=3{
		panic("配置错误请检查此配置 -------> "+ sheetValue)
	}
	infosMap := make(map[string]string)
	s0 := strings.Split(s[0], ":")
	s1 := strings.Split(s[1], ":")
	s2 := strings.Split(s[2], ":")
	if len(s0) !=2 && len(s1)!=2 && len(s2)!=2{
		panic("配置错误请检查此配置 -------> "+ sheetValue)
	}
	var checkParams = map[string]int{"SheetName":0, "KeyID":0, "FindID": 0}
	infosMap[s0[0]] = strings.Replace(s0[1], " ", "", -1)  //strings.Replace(s[0], " ", "")
	checkParams[s0[0]] = 1
	infosMap[s1[0]] = strings.Replace(s1[1], " ", "", -1)
	checkParams[s1[0]] = 1
	infosMap[s2[0]] = strings.Replace(s2[1], " ", "", -1)
	checkParams[s2[0]] = 1
	for _,v := range checkParams{
		if v == 0{
			panic("配置错误请检查此配置 -------> "+ sheetValue)
		}
	}
	return infosMap
}

func CheckPathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}

func StoreWithJson(sheetName string, result interface{})error{
	//将结果存储为json格式文件
	b, err := json.Marshal(result)
	if err != nil {
		fmt.Println("json.Marshal failed:", err)
		return err
	}
	fineName := "./" + sheetName + ".json"
	f,err := os.Create(fineName)
	defer f.Close()
	if err !=nil {
		log.Log("StoreJson",err.Error())
	} else {
		f.Write([]byte(b))
	}
	return nil
}
