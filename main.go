package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"log"
	"os"
	"unicode"

	"github.com/xuri/excelize/v2"
)

const DEF_SHEET_NAME = "Sheet1"
const DEF_SEPARATOR_NAME = ','

var (
	separator               rune
	enclosed                bool
	srcPath, dstPath, sheet string
)

func init() {
	msg := fmt.Sprintf("Usage of %q:\n"+
		"  -s/--src\t\t- source file\n"+
		"  -d/--dst\t\t- destanation file\n"+
		"  -e/--enclosed\t\t- column not enclosed with quotes\n"+
		"  -t/--terminated\t- column's separator\n"+
		"  -n/--name\t\t- sheet name\n"+
		"  -h/--help\t\t- print this\n", os.Args[0])

	for i, arg := range os.Args[1:] {
		switch {
		case arg == "-s" || arg == "--src":
			if i+2 < len(os.Args) {
				srcPath = os.Args[i+2]
			}
		case arg == "-d" || arg == "--dst":
			if i+2 < len(os.Args) {
				dstPath = os.Args[i+2]
			}
		case arg == "-t" || arg == "--terminated":
			if i+2 < len(os.Args) {
				separator = rune(os.Args[i+2][0])
			}
		case arg == "-n" || arg == "--name":
			if i+2 < len(os.Args) {
				dstPath = os.Args[i+2]
			}
		case arg == "-e" || arg == "--enclosed":
			enclosed = true
		case arg == "-h" || arg == "--help":
			fmt.Print(msg)
			os.Exit(0)
		}
	}

	if len(srcPath) == 0 || len(dstPath) == 0 {
		fmt.Print(msg)
		os.Exit(0)
	}

	if len(sheet) == 0 {
		sheet = DEF_SHEET_NAME
	}

	if !unicode.IsPrint(separator) {
		separator = DEF_SEPARATOR_NAME
	}
}

func main() {

	var err error
	var cell string
	var row int
	var record []string

	f, err := os.OpenFile(srcPath, os.O_RDONLY, os.ModePerm)
	if err != nil {
		log.Fatalf("can't open file, error: %v", err)
	}
	defer f.Close()

	reader := csv.NewReader(f)
	reader.Comma = separator
	reader.LazyQuotes = enclosed
	writer := excelize.NewFile()

	for {
		if record, err = reader.Read(); err != nil {
			if err != io.EOF {
				log.Printf("can't read line, got err: %v\n", err)
			}
			break
		} else {
			row++
			if cell, err = excelize.CoordinatesToCellName(1, row); err != nil {
				log.Printf("can't get coordinates of cell, got err: %v\n", err)
			}

			if err = writer.SetSheetRow(sheet, cell, &record); err != nil {
				log.Printf("can't append row to sheet, got err: %v\n", err)
			}
		}
	}

	if err = writer.SaveAs(dstPath); err != nil {
		log.Printf("can't save excel file at: %s, got err: %v\n", dstPath, err)
	}
}
