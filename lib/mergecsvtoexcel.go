package lib

import (
	"encoding/csv"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"io"
	"io/ioutil"
	"log"
	"strconv"
	"strings"
)

type RowHandel struct {}

func SheetAliasConf() map[string]string {
	sheetAlias := make(map[string]string)
	sheetAlias["bank100-fgs-count"] = "百融按分公司统计"
	sheetAlias["bank100-yyzx-detail"] = "百融查询明细(运营中心)"
	sheetAlias["bank100-yyzx-count"] = "百融按运营中心统计"
	sheetAlias["bank100-fgs-detail"] = "百融查询明细(分公司)"
	sheetAlias["tongdun-fgs-detail"] = "同盾查询明细(分公司)"
	sheetAlias["tongdun-fgs-count"] = "同盾统计(分公司)"
	sheetAlias["tongdun-yyzx-detail"] = "同盾查询明细(运营中心)"
	sheetAlias["tongdun-yyzx-count"] = "同盾统计(运营中心)"
	return sheetAlias
}

//
func (r *RowHandel) Handel(rowNum int, rowContent []string, excel *excelize.File, sheet string)  {
	if rowNum >=1  {
		for colIndex, colValue := range rowContent {
			if colValue != "" {
				colName := GetExcelColName(colIndex) + strconv.Itoa(rowNum+2)
				sheetConf := SheetAliasConf()
				sheetName := sheetConf[sheet]
				fmt.Println(sheetName, colIndex, colName,colValue)
				excel.SetCellValue(sheetName, colName, colValue)
			}
		}
	}
}

func CheckRowAndColIndex(sheet string, rowIndex int, colIndex int) bool {
	conf := SheetAliasConf()
	switch sheet {
	case conf["bank100-fgs-count"]:
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "bank100-yyzx-detail":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "bank100-yyzx-count":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "bank100-fgs-detail":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "tongdun-fgs-detail":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "tongdun-fgs-count":
		if rowIndex >=1 && colIndex >=1{
			return true
		}
		break
	case "tongdun-yyzx-detail":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	case "tongdun-yyzx-count":
		if rowIndex >1 && colIndex >1{
			return true
		}
		break
	}
	return false
}

func Merge()  {
	saveFile := "./example_files/credit-all.xlsx"

	//文件
	files := make(map[string]string)
	files["tongdun-fgs-count"] = "./example_files/source.csv"
	//
	sheetConf := SheetAliasConf()
	h := RowHandel{}
	excel,_ := excelize.OpenFile(saveFile)

	for sheet, filename:= range files {

		//fmt.Println(sheetConf[sheet])

		excel.NewSheet(sheetConf[sheet])
		ReadCsv(filename, h, excel, sheet)
	}
	excel.SetActiveSheet(1)
	err := excel.SaveAs(saveFile)

	if err != nil {
		fmt.Println(err)
	}
}

//
func ReadCsv(fileName string, h RowHandel, excel *excelize.File, sheet string) {
	dat, err := ioutil.ReadFile(fileName)
	if err != nil {
		log.Fatal(err)
	}
	rowNum := 0
	r := csv.NewReader(strings.NewReader(string(dat[:])))
	for {
		record, err := r.Read()
		if err == io.EOF {
			break
		}
		if err != nil {
			log.Fatal(err)
		}
		h.Handel(rowNum, record, excel, sheet)
		rowNum++
	}
}

func GetExcelColName(index int) string  {
	data := make(map[int]string)
	letter := []string{"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"}
	for k,v := range letter{
		data[k] = v
	}
	return data[index]
}