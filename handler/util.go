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
