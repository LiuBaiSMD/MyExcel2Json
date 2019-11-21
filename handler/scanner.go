package handler

type sheetInfo1 struct {
	ID int
	BuffID []int
	Pos		string
}
type percentage float32

type sheetInfo struct {
	ID int
	BuffID []int
	Pos		string
	EmptyName []string
	ItemID 	int64
	SkillID []int64
	NumericalRate percentage
	NumericalRates []percentage
	Float32s	[]float32
	Float64s	[]float64
	SingleStruct map[string]interface{}
	SheetInfoMap sheetInfo1
	SheetInfoList []sheetInfo1
}
