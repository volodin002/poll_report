package main

import (
	"fmt"
	"strconv"
	"github.com/tealeg/xlsx"
)

func create_tealeg_report(conf *config, fileName string) {
	
	xlFile := open_tealeg_file(fileName)

	rsh, err := xlFile.AddSheet(REPORT_SHEET)
	if err != nil {
		panic(err.Error())
	}

	sh, ok := xlFile.Sheet[DATA_SHEET]
	if !ok {
		panic("Can not find sheet: Sheet")
	}

	col_map := make(map[int]int)
	colls_targets := make([]int, len(conf.Colls))
	for i, col := range conf.Colls {
		source := colNumber(col.Source)
		target := colNumber(col.Target)
		col_map[source] = target
		colls_targets[i] = target
	}

	confScore := colNumber(conf.Score)
	confStartPoll := colNumber(conf.StartPoll)

	colCnt := confScore
	for _, target := range colls_targets {
		if target > colCnt {
			colCnt = target
		}
	}

	r := 1

	header_style := xlsx.NewStyle()
	font := xlsx.NewFont(11, "Calibri")
	font.Bold = true
	border := xlsx.NewBorder("thin", "thin", "thin", "thin")

	header_style.Font = *font
	header_style.Border = *border
	header_style.ApplyBorder = true

	style := xlsx.NewStyle()
	font = xlsx.NewFont(11, "Calibri")
	style.Font = *font

	for _, row := range sh.Rows {
		newRow := rsh.AddRow()

		// Create all new row cells
		for c, _ := range row.Cells {
			if c == colCnt {
				break
			}
			newRow.AddCell()
		}

		// HEADER
		if r == 1 {

			/*
			for _, targetCell := range newRow.Cells {
				targetCell.SetStyle(header_style)
			}
			*/
			for c, sourceCell := range row.Cells {
				
				tc, ok := col_map[c+1]
				if ok {
					targetCell := newRow.Cells[tc-1]
					val := sourceCell.String()
					targetCell.SetString(val)
					targetCell.SetStyle(header_style)
				} 
			}

			targetCell := newRow.Cells[confScore - 1]
			targetCell.SetString("Сумма балов")
			targetCell.SetStyle(header_style)

			r++
			continue
		}
		

		score := 0
		
		for c, sourceCell := range row.Cells {
			
			if c >= confStartPoll-1 {

				value := sourceCell.String()
				iv, err := strconv.Atoi(value)
				if err != nil {
					cellAxis := cellName(c+1, r)
					println(fmt.Sprintf("Error while convert to int cell %s : %v", cellAxis, err))
					continue
				}
				score = score + iv
				continue
			}

			tc, ok := col_map[c+1]
			if ok {
				targetCell := newRow.Cells[tc-1]
				val := sourceCell.String()
				targetCell.SetString(val)
				targetCell.SetStyle(style)
			} 

		}

		targetCell := newRow.Cells[confScore - 1]
		targetCell.SetInt(score)
		targetCell.SetStyle(style)

		r++
	}

	
	for i, x := range conf.Colls {
		width := x.Width;
		if width == 0 {
			width = 15.0
		}
		target := colls_targets[i]
		rsh.SetColWidth(target, target, width)
	}
	
	width := conf.ScoreWidth
	if width == 0 {
		width = 15.0
	}
	rsh.SetColWidth(confScore, confScore, width)

	xlFile.Save(fileName)
	
}

func open_tealeg_file(fileName string) *xlsx.File {
	xlFile, err := xlsx.OpenFile(fileName)
    if err != nil {
        panic(err.Error())
	}
	
	_, ok := xlFile.Sheet[REPORT_SHEET]
	if ok {
		// Delete Report sheet
		sheets := []*xlsx.Sheet {};
		for _, sheet := range xlFile.Sheets {
			if sheet.Name != REPORT_SHEET {
				sheets = append(sheets, sheet)
			}
		}
		xlFile.Sheets = sheets
		xlFile.Save(fileName)

		xlFile, err = xlsx.OpenFile(fileName)
		if err != nil {
			panic(err.Error())
		}
	}

	return xlFile
}