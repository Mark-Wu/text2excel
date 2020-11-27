package main

import (
	"bufio"
	"fmt"
	"github.com/tealeg/xlsx"
	"io"
	"os"
	"strings"
)


func write_xlsx(sheet *xlsx.Sheet,style *xlsx.Style, word, desc string)  {
	var row  	*xlsx.Row
	var cell 	*xlsx.Cell

	row = sheet.AddRow()
	row.SetHeightCM(1)
	cell = row.AddCell()
	cell.Value = word
	cell = row.AddCell()
	cell.Value = desc
}

type lineString struct {
	value string
	startFlag bool
	endFlag bool
	VoidFlag bool
}


func (str *lineString)formatCheck() bool{
	pos0 := strings.Index(str.value,". ")
	pos1 := strings.Index(str.value,",")

	if pos0 <= 0 {
		str.startFlag = false
	} else if pos1 >=0 && pos1 <= pos0 {
		str.startFlag = false
	} else {
		str.startFlag = true
	}

	if strings.HasSuffix(str.value,".") {
		str.endFlag = true
	} else {
		str.endFlag = false
	}

	str.VoidFlag = false

	if str.startFlag == true && str.endFlag == true {
		return true
	} else {
		return false
	}

}

func (str *lineString)formatCheckAppendix(){
	pos0 := strings.Index(str.value," — ")

	if pos0 <= 0 {
		str.startFlag = false
	} else {
		str.startFlag = true
	}
}

func (str *lineString)ceilFormat(sheet *xlsx.Sheet,style *xlsx.Style) {
	pos0 := strings.Index(str.value,". ")

	newWord := str.value[0:pos0]
	wordDesc := str.value[pos0+1:]
	fmt.Printf("%s ==> %s \n",newWord,wordDesc)
	write_xlsx(sheet,style,newWord,wordDesc)
}

func (str *lineString)ceilFormatAppendix(sheet *xlsx.Sheet,style *xlsx.Style) {
	pos0 := strings.Index(str.value," — ")

	newWord := str.value[0:pos0]
	wordDesc := str.value[pos0+len(" — "):]
	fmt.Printf("%s ==> %s \n",newWord,wordDesc)
	write_xlsx(sheet,style,newWord,wordDesc)
}

func NormalDataTransfer(){

	inputFile, err := os.Open("Dictionary_of_Aeronautical_Terms.txt")
	if err != nil {
		fmt.Printf("Error: %s\n", err)
		return
	}
	defer inputFile.Close()


	var outputFile 	*xlsx.File
	var sheet 	*xlsx.Sheet

	outputFile = xlsx.NewFile()
	sheet, err = outputFile.AddSheet("sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	style := xlsx.NewStyle()

	fill := *xlsx.NewFill("solid", "00FF0000", "FF000000")
	font := *xlsx.NewFont(20, "Verdana")
	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")

	style.Fill = fill
	style.Font = font
	style.Border = border

	style.ApplyFill = true
	style.ApplyFont = true
	style.ApplyBorder = true



	br := bufio.NewReader(inputFile)
	currentCeil := lineString{"",false,false,true}
	lastCeil := lineString{"",false,false,true}
	for {
		line, _, err := br.ReadLine()
		if err == io.EOF {
			break
		}
		if len(line) > 2 {
			if line[0] == 0x0c {
				currentCeil.value = string(line[1:])
			} else {
				currentCeil.value = string(line)
			}
			currentCeil.formatCheck()
			if lastCeil.VoidFlag {
				lastCeil = currentCeil
				continue
			}
			if 	currentCeil.startFlag && lastCeil.endFlag {
				lastCeil.ceilFormat(sheet,style)
				lastCeil = currentCeil
			} else {
				lastCeil.value += " "
				lastCeil.value += currentCeil.value
				lastCeil.endFlag = currentCeil.endFlag
			}
		}

	}
	lastCeil.ceilFormat(sheet,style)


	err = outputFile.Save("Dictionary_of_Aeronautical_Terms.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}

}


func AppendixTransfer(){
	inputFile, err := os.Open("APPENDIX.txt")
	if err != nil {
		fmt.Printf("Error: %s\n", err)
		return
	}
	defer inputFile.Close()

	var outputFile 	*xlsx.File
	var sheet 	*xlsx.Sheet

	outputFile = xlsx.NewFile()
	sheet, err = outputFile.AddSheet("sheet1")
	if err != nil {
		fmt.Printf(err.Error())
	}

	style := xlsx.NewStyle()

	fill := *xlsx.NewFill("solid", "00FF0000", "FF000000")
	font := *xlsx.NewFont(20, "Verdana")
	border := *xlsx.NewBorder("thin", "thin", "thin", "thin")

	style.Fill = fill
	style.Font = font
	style.Border = border

	style.ApplyFill = true
	style.ApplyFont = true
	style.ApplyBorder = true



	br := bufio.NewReader(inputFile)
	currentCeil := lineString{"",false,false,true}

	for {
		line, _, err := br.ReadLine()
		if err == io.EOF {
			break
		}
		if len(line) > 2 {
			if line[0] == 0x0c {
				currentCeil.value = string(line[1:])
			} else {
				currentCeil.value = string(line)
			}
			currentCeil.formatCheckAppendix()

			if 	currentCeil.startFlag {
				currentCeil.ceilFormatAppendix(sheet,style)
			} else {
				continue
			}
		}

	}


	err = outputFile.Save("Apendix.xlsx")
	if err != nil {
		fmt.Printf(err.Error())
	}






}

func main() {
	NormalDataTransfer()
	AppendixTransfer()
}