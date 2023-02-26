package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"

	"strings"

	"github.com/xuri/excelize/v2"
)

var resultPath = "./result"

// 扫描当前文件夹中的文件, 获取xlsx 文件
func scanFolder() []string {

	var arr []string

	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		name := file.Name()
		if path.Ext(name) == ".xlsx" && !strings.Contains(name, "~$") {

			fmt.Println(
				path.Ext(name),
			)
			arr = append(arr, name)
		}
	}
	return arr
}

// 创建结果文件夹
func createResultFolder() {
	_, err := os.Stat(resultPath)
	if !os.IsNotExist(err) {
		err := os.RemoveAll(resultPath)
		if err != nil {
			fmt.Println((err))
		}
	}
	errMk := os.Mkdir(resultPath, os.ModePerm)
	if errMk != nil {
		fmt.Println((errMk))
	}
}

// create excel
func createExcel(fileName string) {
	f := excelize.NewFile()
	// 创建一个工作表
	index, _ := f.NewSheet("Sheet2")
	// 设置单元格的值
	f.SetCellValue("Sheet2", "A2", "Hello world.")
	f.SetCellValue("Sheet1", "B2", 100)
	// 设置工作簿的默认工作表
	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs(path.Join(resultPath, fileName)); err != nil {
		println(err.Error())
	}
}

// read excel
func readExcel(fileName string) {
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		println(err.Error())
		return
	}
	sheets := f.GetSheetList()
	fistSheets := sheets[0]
	println(fistSheets)
	// 获取工作表中指定单元格的值
	// timeCell, err := f.GetCellValue(fistSheets, "日期")
	// if err != nil {
	// 	println(err.Error())
	// 	return
	// }
	// println(timeCell)
	// 获取 Sheet1 上所有单元格
	rows, err := f.GetRows(fistSheets)
	for _, row := range rows {
		for _, colCell := range row {
			print(colCell, "\t")
		}
		println()
	}
}

func main() {
	createResultFolder()
	files := scanFolder()
	// fmt.Println(files)
	//
	for i := 0; i < len(files); i++ {
		readExcel(files[i])
		// createExcel(files[i])
	}
}
