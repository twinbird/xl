package main

import (
	"bufio"
	"flag"
	"fmt"
	"github.com/tealeg/xlsx"
	"math"
	"os"
	"strconv"
	"strings"
)

func main() {
	flag.Usage = func() {
		fmt.Fprintf(os.Stderr, `
  xl: Row oriented Excel operation tool

  Usage:
      xl -in [filepath]
      or
      xl -out [filepath]

  Option:
`)
		flag.PrintDefaults()
	}
	in := flag.String("in", "", "input excel file path")
	out := flag.String("out", "", "output excel file path")
	sep := flag.String("s", " ", "text column separater character")
	sheet := flag.String("S", "", "operation sheet name")
	base := flag.String("b", "", "base cell position")

	flag.Parse()

	if (*in == "" && *out == "") || (*in != "" && *out != "") {
		flag.Usage()
	}

	cellPos, err := optToCellPos(*base)
	if err != nil {
		fmt.Fprintf(os.Stderr, err.Error())
		os.Exit(1)
	}

	if *in != "" {
		if err := excelToStdout(*in, *sep, *sheet, cellPos); err != nil {
			fmt.Fprint(os.Stderr, err.Error())
			os.Exit(1)
		}
	}
	if *out != "" {
		if err := StdinToExcel(*out, *sep, *sheet, cellPos); err != nil {
			fmt.Fprint(os.Stderr, err.Error())
			os.Exit(1)
		}
	}
}

type CellPos struct {
	Column int
	Row    int
}

func optToCellPos(opt string) (*CellPos, error) {
	if opt == "" {
		return &CellPos{}, nil
	}
	opts := strings.Split(opt, ",")
	if len(opts) != 2 {
		return nil, fmt.Errorf("Invalid option: %s", opt)
	}
	col, err := colIdxToI(opts[0])
	if err != nil {
		return nil, err
	}
	row, err := strconv.Atoi(opts[1])
	if err != nil {
		return nil, err
	}
	return &CellPos{col, row - 1}, nil
}

// if necessary convert A1 Format column index to R1C1 Format index.
func colIdxToI(idxChar string) (int, error) {
	c := strings.ToUpper(idxChar)
	d, err := strconv.Atoi(c)
	if err == nil {
		return d, nil
	}
	ret := 0
	for i, r := range c {
		if r < 'A' || 'Z' < r {
			return 0, fmt.Errorf("%s is not a alphabet", r)
		}
		ret += int(math.Pow(float64(26), float64(len(c)-i-1))) * (int(r) - int('A') + 1)
	}
	return ret - 1, nil
}

func excelToStdout(inFilePath string, sep string, sheet string, cellPos *CellPos) error {
	f, err := xlsx.OpenFile(inFilePath)
	if err != nil {
		return err
	}

	s := f.Sheets[0]
	if sheet != "" {
		s = f.Sheet[sheet]
	}
	if s == nil {
		return fmt.Errorf("There is no sheet of '%s'", sheet)
	}
	for ridx, row := range s.Rows {
		if ridx < cellPos.Row {
			continue
		}
		fields := make([]string, 0)
		for cidx, cell := range row.Cells {
			if cidx < cellPos.Column {
				continue
			}
			s := cell.String()
			if err != nil {
				return err
			}
			if s == "" {
				break
			}
			fields = append(fields, s)
		}
		if len(fields) == 0 {
			break
		}
		outStr := strings.Join(fields, sep)
		fmt.Println(outStr)
	}
	return nil
}

func isExist(filename string) bool {
	_, err := os.Stat(filename)
	return err == nil
}

func StdinToExcel(outFilePath string, sep string, sheet string, cellPos *CellPos) error {
	var f *xlsx.File
	var sh *xlsx.Sheet
	if isExist(outFilePath) == false {
		var err error
		f = xlsx.NewFile()
		sh, err = f.AddSheet("sheet1")
		if err != nil {
			return err
		}
	} else {
		var err error
		f, err = xlsx.OpenFile(outFilePath)
		if err != nil {
			return err
		}
		sh = f.Sheets[0]
		if sheet != "" {
			sh = f.Sheet[sheet]
		}
		if sh == nil {
			return fmt.Errorf("There is no sheet of '%s'", sheet)
		}
	}

	rowIdx := cellPos.Row
	for i := 0; i < cellPos.Row; i++ {
		sh.AddRow()
	}
	scanner := bufio.NewScanner(os.Stdin)
	for scanner.Scan() {
		s := scanner.Text()
		strs := strings.Split(s, sep)
		for i, v := range strs {
			var row *xlsx.Row
			if len(sh.Rows) <= rowIdx {
				row = sh.AddRow()
			} else {
				row = sh.Rows[rowIdx]
			}
			var cell *xlsx.Cell
			for len(row.Cells) <= i + cellPos.Column {
				cell = row.AddCell()
			}
			cell = row.Cells[i+cellPos.Column]
			cell.Value = v
		}
		rowIdx++
	}
	if err := scanner.Err(); err != nil {
		return err
	}
	if err := f.Save(outFilePath); err != nil {
		return err
	}
	return nil
}
