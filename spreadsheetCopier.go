package main

import (
	"bufio"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"os"
	"strconv"
	"strings"
)

func main() {
	fmt.Print("Enter the name of the file to be copied: ")
	reader := bufio.NewReader(os.Stdin)
	fileInput, err := reader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print("Enter the name of the Sheet you would like to copy: ")
	newReader := bufio.NewReader(os.Stdin)
	sheetInput, err := newReader.ReadString('\n')
	fileInput = strings.TrimSpace(fileInput)
	sheetInput = strings.TrimSpace(sheetInput)

	originalFile, err := excelize.OpenFile(fileInput)
	if err != nil {
		fmt.Println(err)
		return
	}
	newFile := excelize.NewFile()
	rows, err := originalFile.GetRows(sheetInput)
	i := 0
	for _, row := range rows {
		for _, colCell := range row {
			i++
			cell := strconv.Itoa(i)
			cell = "A" + cell
			newFile.SetCellValue("Sheet1", cell, colCell)

		}
	}
	if err := newFile.SaveAs("Book2.xlsx"); err != nil {
		fmt.Println(err)
	}
}
