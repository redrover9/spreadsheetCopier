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
	fileReader := bufio.NewReader(os.Stdin)
	fileInput, err := fileReader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print("Enter the name of the Sheet you would like to copy: ")
	sheetReader := bufio.NewReader(os.Stdin)
	sheetInput, err := sheetReader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}
	fmt.Print("Starting from which row would you like to copy? ")
	rowReader := bufio.NewReader(os.Stdin)
	rowInput, err := rowReader.ReadString('\n')
	if err != nil {
		log.Fatal(err)
	}

	fileInput = strings.TrimSpace(fileInput)
	sheetInput = strings.TrimSpace(sheetInput)
	rowInput = strings.TrimSpace(rowInput)
	rowNum, err := strconv.Atoi(rowInput)
	rowNum--
	originalFile, err := excelize.OpenFile(fileInput)
	if err != nil {
		fmt.Println(err)
		return
	}
	newFile := excelize.NewFile()
	rows, err := originalFile.GetRows(sheetInput)
	for _, row := range rows {
		for _, colCell := range row {
			rowNum++
			cell := strconv.Itoa(rowNum)
			cell = "A" + cell
			newFile.SetCellValue("Sheet1", cell, colCell)

		}
	}
	if err := newFile.SaveAs("Book2.xlsx"); err != nil {
		fmt.Println(err)
	}
}
