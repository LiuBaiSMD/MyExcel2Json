package handler


import (
"fmt"
"github.com/tealeg/xlsx"
"strings"
"strconv"
"encoding/json"
"errors"
"os"
)

type sheetInfo1 struct {
	ID int
}

type sheetInfo struct {
	ID int
	SheetInfoMap sheetInfo1
	SheetInfoList []sheetInfo1
}

func Change() {
	test := `{"test0":1, "test1":2, "test3":{"test4":"4"}, "test5":[1,2,3,4], "test6":"testStr"}`
	getStrMap("", test)
	excelFileName := "1.xlsx"
	var xlFile *xlsx.File
	var err error
	xlFile, err = xlsx.OpenFile(excelFileName)
	if err != nil {
		panic(err)
	}

	//测试区
	testThis(xlFile)

	for _, sheet := range xlFile.Sheets {

		for i:= 0;i<len(sheet.Rows);i++{
			for j := 0 ; j<len(sheet.Rows[i].Cells);j++{
				//fmt.Println("输出： ", sheet.Rows[i].Cells[j].String())
			}
		}

		//第一个先把所有数据的Key得到
		if strings.HasPrefix(sheet.Name, "@"){
			fmt.Println("=============此表为子表不做处理")
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
				//开始构造结果体
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
			if ifAdd{
				sheetResult = append(sheetResult, rowRessult)
				var test sheetInfo
				jsonStr, _ := json.Marshal(rowRessult)
				json.Unmarshal([]byte(jsonStr), &test)
				fmt.Println("struct : -------> ", test)
			}
		}
		for _, j := range sheetResult{
			fmt.Println("sheetResult: ", j)
		}
		b, err := json.Marshal(sheetResult)
		if err != nil {
			fmt.Println("json.Marshal failed:", err)
			return
		}

		f,err := os.Create("./test.json")
		defer f.Close()
		if err !=nil {
			fmt.Println(err.Error())
		} else {
			f.Write([]byte(b))

		}
		fmt.Println("b:", string(b))
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
			//fmt.Println("strList", strList)
			return nil, err
		}
		intList = append(intList,intStr)
	}
	//fmt.Println("intList: ", intList)
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
			//fmt.Println("strList", strList)
			return nil, err
		}
		intList = append(intList,intStr)
	}
	//fmt.Println("intList: ", intList)
	return intList, nil
}

func getInt(decs, value string)(int, error){
	dft := 0
	if value==""{
		return dft, nil
	}
	i, err := strconv.Atoi(value)
	if err != nil{
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
		return dft, err
	}
	return i, nil
}

func getStrMap(decs, value string)(map[string]interface{}, error){
	var mapResult map[string]interface{}
	err:= json.Unmarshal([]byte(value), &mapResult)
	fmt.Println("map: ", mapResult)
	return mapResult,err
}

func getStrList(desc, value string)([]string, error){
	spliter := "|"
	StrList := strings.Split(value, spliter)
	return StrList, nil
}

func getOtherSheetInfo(Xfile *xlsx.File, keyID, valType, decs, sheetName, findID string)(interface{}, error){
	//keyID为唯一识别的ID标识，如"ID" 或者"id"，等索引的标识服务，excel第一行中的数据
	fmt.Println(keyID, sheetName, findID)
	var result interface{}
	var err error
	if valType == "[]SheetInfo"{
		fmt.Println("fine List")
		result, err = getOtherSheetInfoList(Xfile, keyID, decs, sheetName, findID )
	}
	if valType == "SheetInfo"{
		fmt.Println("fine Map")
		result, err = getOtherSheetInfoMap(Xfile, keyID, decs, sheetName, findID )
	}
	return result, err
}

func getOtherSheetInfoList(Xfile *xlsx.File, keyID, decs, sheetName, findID string)(interface{}, error){
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
		fmt.Println("if this: ", keyID)
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
			fmt.Println("找到啦： ", i)
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

func getOtherSheetInfoMap(Xfile *xlsx.File, keyID, decs, sheetName, findID string)(interface{}, error){
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
	var findIDResult interface{}
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
	if checkIfNilStr(cellStrValue){
		value = ""
		return value, nil
	}
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
		value, err = getStrMap(desc, cellStrValue)
	case "[]string":
		value, err = getStrList(desc, cellStrValue)
	case "[]SheetInfo":
		infosMap := splitSheetInfo(cellStrValue)
		value, err = getOtherSheetInfo(Xfile, infosMap["KeyID"], "[]SheetInfo", desc, infosMap["SheetName"], infosMap["FindID"])
		fmt.Println("testinfosList: ", infosMap, value, err)
		//value, err = getOtherSheetInfo()
	case "SheetInfo":
		infosMap := splitSheetInfo(cellStrValue)
		value, err = getOtherSheetInfo(Xfile, infosMap["KeyID"], "SheetInfo", desc, infosMap["SheetName"], infosMap["FindID"])
		fmt.Println("testinfosMap: ", infosMap, value, err)
		//value, err = getOtherSheetInfo()
	default:
		value = cellStrValue
	}
	return value, err
}

func testThis(Xfile *xlsx.File){
	t := "[]SheetInfo"
	result, err :=getOtherSheetInfo(Xfile, "ID", t, "","Sample", "101")
	if t == "[]SheetInfo"{
		for i, j:= range result.([]interface{}){
			fmt.Println(i, " : ", j, err)
		}
	}
	if t == "SheetInfo"{
		fmt.Println("map : ", result, err)
	}
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
	infosMap[s0[0]] = s0[1]
	checkParams[s0[0]] = 1
	infosMap[s1[0]] = s1[1]
	checkParams[s1[0]] = 1
	infosMap[s2[0]] = s2[1]
	checkParams[s2[0]] = 1
	for _,v := range checkParams{
		if v == 0{
			panic("配置错误请检查此配置 -------> "+ sheetValue)
		}
	}
	return infosMap
}

