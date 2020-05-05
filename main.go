package main

import (
	"html/template"
	"log"
	"net/http"
	"strconv"

	"github.com/Luxurioust/excelize"
)

func ParseHTMLTemplate(w http.ResponseWriter, r *http.Request) {
	// 获取html模板
	htmlTemplate, err := template.ParseFiles("./index.html")
	if err != nil {
		log.Fatal(err)
	}
	tableContent := ParseExcelFile("./test.xlsx")
	htmlTemplate.Execute(w, template.HTML(tableContent))
}
func ParseExcelFile(filename string) string {
	excelFile, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatal(err)
	}
	var htmlContent = ""
	sheets := excelFile.GetSheetList()
	// fmt.Println(sheets)
	for _, val := range sheets {
		htmlContent += "<table class='parseTable' data-sheet='" + val + "' border='0' cellspacing='0' cellpadding='0'>"
		rows, err := excelFile.GetRows(val)
		if err != nil {
			log.Fatal(err)
		}
		if len(rows) > 0 {
			htmlContent += addTableHead(len(rows[0]), val, excelFile)
		}
		htmlContent += "<tbody>"
		for index, row := range rows {
			htmlContent += "<tr><td style='width: 50px;text-align: center;font-weight: bold;'>" + strconv.Itoa(index+1) + "</td>"
			rowHeight, err := excelFile.GetRowHeight(val, index)
			if err != nil {

			}
			// fmt.Println("行高", rowHeight)
			for colIndex, value := range row {
				colWidth, err := excelFile.GetColWidth(val, getColName(colIndex))
				// fmt.Println("输出列宽--", colWidth)
				if err != nil {

				}
				htmlContent += "<td style='display: block;width:" + strconv.FormatFloat(colWidth*10, 'f', -1, 64) + "px;height:" + strconv.FormatFloat(rowHeight, 'f', -1, 64) + "px'/>" + value + "</td>"
				// fmt.Printf("\t%s", value)
			}
			// fmt.Println()
			htmlContent += "</tr>"
		}
		htmlContent += "</tbody></table>"
	}
	return htmlContent
}
func addTableHead(len int, sheetName string, excelFile *excelize.File) string {
	headContent := "<thead><tr><th style='width: 50px;height: 30px;'></th>"
	for i := 0; i < len; i++ {
		colWidth, err := excelFile.GetColWidth(sheetName, getColName(i))
		if err != nil {
		}
		headContent += "<th  style='display: block;width:" + strconv.FormatFloat(colWidth*10, 'f', -1, 64) + "px;height:30px'>" + getColName(i) + "</th>"
	}
	headContent += "</tr></thead>"
	// fmt.Println(headContent)
	return headContent
}
func getColName(i int) string {
	var letterA byte = 'A'
	letterIndex := 0
	letterSecond := ""
	if i < 26 {
		letterIndex = i
	} else {
		letterIndex = i / 26
	}
	letter := string(byte(int(letterA) + letterIndex))
	if i >= 26 {
		letterSecondIndex := i % 26
		letterSecond = string(byte(int(letterA) + letterSecondIndex))
	}
	return letter + letterSecond
}
func main() {
	server := http.Server{
		Addr: "127.0.0.1:8080",
	}
	log.Println("服务运行：127.0.0.1：8080")
	http.HandleFunc("/table", ParseHTMLTemplate)
	server.ListenAndServe()
}
