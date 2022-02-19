package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

var dirParam string
var pwd string
var param = flag.String("dir", "./", "配置文件路径")

func main() {
	flag.Parse()
	dirParam, _ = filepath.Abs(*param)
	pwd, _ = os.Getwd()

	// 遍历目录
	err := filepath.Walk(*param, walkFunc)
	if err != nil {
		fmt.Println(err)
	}
}

// 对遍历的文件执行的操作
func walkFunc(path string, info os.FileInfo, err error) error {
	// 目录
	if info.IsDir() {
		return nil
	}

	// 编制中临时文件
	if filepath.HasPrefix(filepath.Base(path), "~$") {
		return nil
	}

	// 扩展名
	ext := filepath.Ext(path)
	if ext == ".xlsx" || ext == ".xls" {
		return parseExcel(path)
	}

	return nil
}

// WalkDir 遍历目录获取指定后缀的文件名列表
func WalkDir(filePath string, suffix []string) ([]string, error) {
	fs := []string{}
	files, err := ioutil.ReadDir(filePath)
	if err != nil {
		return nil, err
	}
	for _, v := range files {
		if v.IsDir() {
			WalkDir(v.Name(), suffix)
		} else {
			fs = append(fs, v.Name())
		}
	}
	return fs, nil
}

// 解析Excel表
func parseExcel(xlsxName string) error {
	f, err := excelize.OpenFile(xlsxName)
	if err != nil {
		return err
	}
	defer f.Close()

	// 获取配置表名
	fileName, err := f.GetCellValue("Sheet1", "B1")
	if err != nil {
		return fmt.Errorf("[%s]: %s", xlsxName, err.Error())
	}

	cfgType, err := f.GetCellValue("Sheet1", "B2")
	if err != nil {
		return fmt.Errorf("[%s]: %s", xlsxName, err.Error())
	}

	var cfgs interface{}
	if cfgType == "列表" {
		cfgs, err = parseArray(f)
	} else if cfgType == "单项" {
		cfgs, err = parseSingle(f)
	} else {
		return fmt.Errorf("[%s]: 配置类型错误，仅支持“列表”与“单项”", xlsxName)
	}

	if err != nil {
		return fmt.Errorf("[%s]: %s", xlsxName, err.Error())
	}

	// 创建目录
	absPath, _ := filepath.Abs(filepath.Dir(xlsxName))
	relativePath := strings.TrimPrefix(absPath, dirParam)
	jsonFilePath := filepath.Join(pwd, "json", relativePath, fileName+".json")
	err = os.MkdirAll(filepath.Dir(jsonFilePath), 0755)
	if err != nil {
		return err
	}

	// 创建json文件
	openFlag := os.O_CREATE | os.O_WRONLY | os.O_TRUNC
	fd, err := os.OpenFile(jsonFilePath, openFlag, os.FileMode(0644))
	if err != nil {
		return fmt.Errorf("[%s]: %s", xlsxName, err.Error())
	}

	d, _ := json.MarshalIndent(cfgs, "", "\t")
	fd.Write(d)
	fd.Close()
	return nil
}

// 解析列表配置项
func parseArray(f *excelize.File) (interface{}, error) {
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, err
	}
	fields := rows[3] // 字段名
	fieldLen := len(fields)
	types := rows[4] // 字段类型
	typeLen := len(types)
	if fieldLen != typeLen {
		return nil, fmt.Errorf("字段名与字段类型对不齐")
	}

	cfgs := []interface{}{}
	for idx, row := range rows[5:] {
		if len(row) != fieldLen {
			return nil, fmt.Errorf("在第%d行%d列配置项与字段对不齐", idx+6, len(row)+1)
		}
		record := map[string]interface{}{}
		for i := 0; i < fieldLen; i++ {
			value, err := conv(row[i], types[i])
			if err != nil {
				return nil, fmt.Errorf("在第%d行%d列错误: %s", idx+6, i+1, err.Error())
			}
			record[fields[i]] = value
		}
		cfgs = append(cfgs, record)
	}
	return cfgs, nil

}

// 解析单项配置
func parseSingle(f *excelize.File) (interface{}, error) {
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		return nil, err
	}

	record := map[string]interface{}{}
	for idx, row := range rows[3:] {
		if len(row) < 3 {
			return nil, fmt.Errorf("在第%d行配置项对不齐", idx+4)
		}

		value, err := conv(row[2], row[1])
		if err != nil {
			return nil, fmt.Errorf("在第%d行错误: %s", idx+4, err.Error())
		}
		record[row[0]] = value
	}
	return record, nil
}

func conv(value, typ string) (interface{}, error) {
	switch typ {
	case "int":
		i, err := strconv.Atoi(value)
		return i, err
	case "float":
		f, err := strconv.ParseFloat(value, 64)
		return f, err
	case "string":
		return value, nil
	}

	return value, nil
}
