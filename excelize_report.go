package main

/*
import (
	"fmt"
	"strconv"
	"github.com/360EntSecGroup-Skylar/excelize"
)


func create_excelize_report(conf *config, fileName string) {

	//
	f, err := excelize.OpenFile(fileName)
	if err != nil {
		panic(err.Error())
	}

	sheets := f.GetSheetMap()

	for _, name := range sheets {
		println(name)
		
		if name == "Report" {
			f.DeleteSheet(name)
		}
	}

	f.NewSheet("Report")

	err = render_excelize_report(conf, f)
	if err !=nil {
		println(err.Error())

		return
	}

	f.Save()
}

func render_excelize_report(conf *config, f *excelize.File) error {
	col_map := make(map[int]int)
	for _, col := range conf.Colls {
		col_map[col.Source] = col.Target
	}

	// Header
	c := 1
	r := 1
	for {
		cellAxis := cellName(c, r)
		value, _ := f.GetCellValue("Sheet", cellAxis)
		if value == "" {
			break
		}
		tc, ok := col_map[c]
		if ok {
			cellAxis = cellName(tc, r)
			f.SetCellValue("Report", cellAxis, value)
		} else if conf.Score == c {
			cellAxis = cellName(c, r)
			f.SetCellValue("Report", cellAxis, "Сумма балов")
		}
		c++
	}

	columnsCnt := c
	r++

	for {
		score := 0
		for c = 1 ; c < columnsCnt; c++ {

			if c >= conf.StartPoll {
				cellAxis := cellName(c, r)
				value, err := f.GetCellValue("Sheet", cellAxis)
				if err != nil {
					return err
				}
				if value == "" {
					return nil
				}
				iv, err := strconv.Atoi(value)
				if err != nil {
					println(fmt.Sprintf("Error while convert to int cell %s : %v", cellAxis, err))
					continue
				}
				score = score + iv
				continue
			}

			tc, ok := col_map[c]
			if ok {
				cellAxis := cellName(c, r)
				value, err := f.GetCellValue("Sheet", cellAxis)
				if err != nil {
					return err
				}
				if value == "" {
					continue
				}

				cellAxis = cellName(tc, r)
				f.SetCellValue("Report", cellAxis, value)
			}  
		}

		cellAxis := cellName(conf.Score, r)
		f.SetCellValue("Report", cellAxis, score)

		println(fmt.Sprintf("Processed row: %d", r) )

		r++
	}
}
*/