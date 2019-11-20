package handler

import (
	"strings"
)

func checkIfNilStr(s string)bool{
	s = strings.Replace(s, " ", "", -1)
	s = strings.Replace(s, "\n", "", -1)
	if s == ""{
		return true
	}
	return false
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

func getDefaultVal(valtype string)interface{}{
	//对于空值设置初始值
	valtype = strings.Replace(valtype, " ", "", -1)
	var value interface{}
	switch valtype {
	case "[]int32":
		value = 0
	case "[]int64":
		value = []int64{}
	case "int32":
		value = 0
	case "int64":
		value = (int64)(0)
	case "float64":
		value = (float64)(0)
	case "float32":
		value = (float32)(0)
	case "[]float64":
		value = []float64{}
	case "[]float32":
		value = []float32{}
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