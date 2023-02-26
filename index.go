package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
)

// 扫描当前文件夹中的文件, 获取xlsx 文件
func scanFolder() []string {

	var arr []string

	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatal(err)
	}

	for _, file := range files {
		if path.Ext(file.Name()) == ".xlsx" {

			fmt.Println(
				path.Ext(file.Name()),
			)
			arr = append(arr, file.Name())
		}
	}
	return arr
}

// 创建结果文件夹
func createResultFolder() {
	var path = "./result"
	_, err := os.Stat(path)
	if !os.IsNotExist(err) {
		err := os.RemoveAll(path)
		if err != nil {
			fmt.Println((err))
		}
	}
	errMk := os.Mkdir(path, os.ModePerm)
	if errMk != nil {
		fmt.Println((errMk))
	}
}

func main() {
	createResultFolder()
	files := scanFolder()
	fmt.Println(files)
}
