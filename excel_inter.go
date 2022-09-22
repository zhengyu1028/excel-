package main

import (
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"io/ioutil"
	"log"
	"os"
	"strconv"
	"strings"
)

type inte interface {
	getexecpath(pat string) ([]string, error)                                            //获取表
	Getrows(string) ([][]string, error)                                                  //获取表得所有单元格得值
	WriteRows(rows [][]string, init_index *int, f *excelize.File, filename string) error //写入新的excel
}

type info struct {
	Init_Index int    //定义初始下标
	ChoosePath string //创建表得名称
}

func (i info) getexecpath(pat string) ([]string, error) { //获取目录下所有execl表
	files, err := ioutil.ReadDir(pat)
	if err != nil {
		log.Fatalln(err)
	}
	file_list := make([]string, 0)
	for _, f := range files {
		//if f.IsDir() {
		//	fpath := pat + "\\" + f.Name()
		//	L, _ := i.getexecpath(fpath)
		//	file_list = append(file_list, L...)
		//} else {
		fpath := pat + "\\" + f.Name()
		file_list = append(file_list, fpath)
	}

	return file_list, nil
}

func (i info) Getrows(filename string) ([][]string, error) {
	f, err := excelize.OpenFile(filename) //打开excel
	if err != nil {
		return nil, errors.New("文件不存在或文件受损")
	}
	sheetlist := f.GetSheetList()   // 获取excel所有工作表
	for _, ele := range sheetlist { // 遍历工作表
		if strings.Contains(ele, "网络设备") || strings.Contains(filename, string([]rune(ele)[:2])) || strings.Contains(filename, string([]rune(ele)[:3])) || strings.Contains(filename, ele) {
			rows, err := f.GetRows(ele) //获取所有单元格的值
			if err != nil {
				return nil, errors.New("获取所有单元格失败")

			}
			return rows, nil

		}
	}
	return nil, err

}

func (i info) WriteRows(rows [][]string, init_index *int, f *excelize.File, filename string) error {
	index_list := make([]int, 11)

	//下标加减
	//func() {
	//	if err := recover(); err != nil {
	//		fmt.Println(err)
	//	}
	//}()
	for indexx, ele := range rows { //遍历二维切片
		*init_index += 1
		fmt.Println(*init_index)
		bt := "A" + strconv.Itoa(*init_index-1)
		if indexx == 0 {
			for ind, element := range ele { //遍历切片并标记下标

				switch {
				case strings.Contains(element, "设备标签") || strings.Contains(element, "设备编码"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "设备序列号"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "设备类型"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "机房名称"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "机房地址"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "项目名称"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "品牌"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "型号"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "交换机层级"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "网管网IP"):
					index_list = append(index_list, ind)
				case strings.Contains(element, "网管网网关"):
					index_list = append(index_list, ind)

				}

			}
			f.SetCellValue("sheet2", bt, filename)
			f.NewStyle("#666666")
		} else if indexx == len(rows)-1 {
			*init_index += 1
		}

		//s := indexx + 2
		sheetindex := "A" + strconv.Itoa(*init_index) //设置单元格
		sheetindex1 := "B" + strconv.Itoa(*init_index)
		sheetindex2 := "C" + strconv.Itoa(*init_index)
		sheetindex3 := "D" + strconv.Itoa(*init_index)
		sheetindex4 := "E" + strconv.Itoa(*init_index)
		sheetindex5 := "F" + strconv.Itoa(*init_index)
		sheetindex6 := "G" + strconv.Itoa(*init_index)
		sheetindex7 := "H" + strconv.Itoa(*init_index)
		sheetindex8 := "I" + strconv.Itoa(*init_index)
		sheetindex9 := "J" + strconv.Itoa(*init_index)
		sheetindex10 := "K" + strconv.Itoa(*init_index)

		//
		if len(ele) >= index_list[21] {
			f.SetCellValue("sheet2", sheetindex, ele[index_list[11]])

			f.SetCellValue("sheet2", sheetindex1, ele[index_list[12]])

			f.SetCellValue("sheet2", sheetindex2, ele[index_list[13]])

			f.SetCellValue("sheet2", sheetindex3, ele[index_list[14]])

			f.SetCellValue("sheet2", sheetindex4, ele[index_list[15]])

			f.SetCellValue("sheet2", sheetindex5, ele[index_list[16]])

			f.SetCellValue("sheet2", sheetindex6, ele[index_list[17]])

			f.SetCellValue("sheet2", sheetindex7, ele[index_list[18]])

			f.SetCellValue("sheet2", sheetindex8, ele[index_list[19]])

			f.SetCellValue("sheet2", sheetindex9, ele[index_list[20]])

			f.SetCellValue("sheet2", sheetindex10, ele[index_list[21]])
			//	//break
			//
		} else {
			f.SetCellValue("sheet2", sheetindex, "此行为空，请对照原表核对")
		}
	}

	return nil
}

func main() {
	var i info = info{ //声明初始信息
		Init_Index: 1,
	}
	var intergration inte //声明接口

	intergration = i
	addr, _ := os.Getwd() //获取当前工作目录

	e_list, err := intergration.getexecpath(addr) //获取目录文件列表

	if err != nil {
		log.Fatalln("文件路径不正确")
	}
	f := excelize.NewFile()
	index := f.NewSheet("sheet2")
	err = f.SetColWidth("Sheet2", "A", "K", 20)
	if err != nil {
		fmt.Println(err)
	}
	f.SetActiveSheet(index)
	if err := f.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

	s, err := excelize.OpenFile("Book1.xlsx")
	for _, ele := range e_list { //循环
		if strings.Contains(ele, "网络设备") && strings.Contains(ele, ".xlsx") {
			rows, err := intergration.Getrows(ele)
			if err != nil {
				log.Println("文件打开失败，excel表格不存在或权限不足或文件受损")
				log.Println(err)
				continue
			}

			err = intergration.WriteRows(rows, &i.Init_Index, s, ele)
		} else {
			continue
		}
	}
	indexx := s.NewSheet("sheet2")
	s.SetActiveSheet(indexx)
	if err := s.SaveAs("Book1.xlsx"); err != nil {
		fmt.Println(err)
	}

}
