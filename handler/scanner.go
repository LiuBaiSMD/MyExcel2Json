package handler

type sheetInfo1 struct {
	ID int
	BuffID []int
	Pos		string
}

type sheetInfo struct {
	ID int
	BuffID []int
	Pos		string
	ItemID int
	NumericalRate float32
	SheetInfoMap sheetInfo1
	SheetInfoList []sheetInfo1
}
