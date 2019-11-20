package handler

type sheetInfo1 struct {
	ID int
}

type sheetInfo struct {
	ID int
	BuffID []int
	Pos		string
	SheetInfoMap sheetInfo1
	SheetInfoList []sheetInfo1
}
