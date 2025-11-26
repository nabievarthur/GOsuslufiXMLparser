package service

import (
	"encoding/xml"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ===== СТРУКТУРЫ =====

type XmlParser struct{}

func NewXMLParser() *XmlParser {
	return &XmlParser{}
}

type List struct {
	XMLName  xml.Name   `xml:"List"`
	Document []Document `xml:"Document"`
}

type Document struct {
	RequestInfo RequestInfo `xml:"RequestInfo"`
	DocNumber   string      `xml:"DocumentID"`
}

type RequestInfo struct {
	ConvictionPerson ConvictionPerson `xml:"ConvictionPerson"`
}

type ConvictionPerson struct {
	CPSurname    string   `xml:"CPSurname"`
	CPName       string   `xml:"CPName"`
	CPPatronymic string   `xml:"CPPatronymic"`
	CPBirthday   string   `xml:"CPBirthday"`
	CPLastFIO    *LastFIO `xml:"CPLastFIO"`
}

type LastFIO struct {
	CPLSurname string `xml:"CPLSurname"`
}

type XLSRow struct {
	Surname         string // Фамилия
	Name            string // Имя
	Patronymic      string // Отчество
	BirthYear       string // Год рождения
	BirthMonth      string // Месяц рождения
	BirthDay        string // День рождения
	Result          string // Результат
	WantedPersons   string // Розыск лиц
	OSKRegion       string // ОСК регион
	OSKGIAZ         string // ОСК ГИАЦ
	AdminPracticeR  string // Адмпрактика регион
	AdminPracticeF  string // Адмпрактика ФИС-М
	ZAGSDeath       string // ЗАГС рег.смерти
	Restricted      string // Запретники
	PassportRF      string // Паспорт РФ
	DeportationMode string // Реж.высылки
	DocumentNumber  string // № документа (из XML)
}

// меняем дату

func convertDate(d string) string {
	parts := strings.Split(d, ".")
	if len(parts) != 3 {
		return ""
	}
	return fmt.Sprintf("%s;%s;%s", parts[2], parts[1], parts[0])
}

// ===== ПАРСИНГ XML =====

func ParseXMLToLines(xmlData []byte) ([]string, error) {
	var list List
	err := xml.Unmarshal(xmlData, &list)
	if err != nil {
		return nil, err
	}

	var lines []string

	for _, doc := range list.Document {
		p := doc.RequestInfo.ConvictionPerson
		date := convertDate(p.CPBirthday)

		// текущая фамилия
		lines = append(lines,
			fmt.Sprintf("%s;%s;%s;%s",
				p.CPSurname, p.CPName, p.CPPatronymic, date))

		// старая фамилия (если есть)
		if p.CPLastFIO != nil && p.CPLastFIO.CPLSurname != "" {
			lines = append(lines,
				fmt.Sprintf("%s;%s;%s;%s",
					p.CPLastFIO.CPLSurname, p.CPName, p.CPPatronymic, date))
		}
	}

	return lines, nil
}

func (x *XmlParser) ParseXMLToFile(filename string) (string, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return "", err
	}

	lines, err := ParseXMLToLines(data)
	if err != nil {
		return "", err
	}

	var result strings.Builder
	for _, line := range lines {
		result.WriteString(line)
		result.WriteString("\n")
	}
	return result.String(), nil
}

func ReadXLSFile(filename string) ([]XLSRow, error) {
	// пробуем как excel
	f, err := excelize.OpenFile(filename)
	if err != nil {
		fmt.Printf("Не удалось открыть как Excel, пробуем как HTML: %v\n", err)
		// Если не получилось, пробуем прочитать как HTML
		return readHTMLTable(filename)
	}
	defer f.Close()

	// Логика для настоящих Excel файлов
	rows, err := f.GetRows("Sheet1")
	if err != nil {
		sheets := f.GetSheetList()
		if len(sheets) > 0 {
			rows, err = f.GetRows(sheets[0])
			if err != nil {
				return nil, err
			}
		} else {
			return nil, fmt.Errorf("нет доступных листов в файле")
		}
	}

	return parseRowsToXLSRows(rows), nil
}

// читаем  HTML-таблицу из файла
func readHTMLTable(filename string) ([]XLSRow, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}

	content := string(data)

	trRegex := regexp.MustCompile(`(?is)<tr[^>]*>(.*?)</tr>`)
	tdRegex := regexp.MustCompile(`(?is)<t[dh][^>]*>(.*?)</t[dh]>`)

	matches := trRegex.FindAllStringSubmatch(content, -1)
	if matches == nil {
		return nil, fmt.Errorf("таблица <tr>...</tr> не найдена в HTML")
	}

	var rows [][]string

	for _, trMatch := range matches {
		trContent := trMatch[1]

		var cells []string

		tds := tdRegex.FindAllStringSubmatch(trContent, -1)
		for _, td := range tds {
			text := stripHTML(td[1])
			text = strings.ReplaceAll(text, "&nbsp;", " ")
			text = strings.TrimSpace(text)
			cells = append(cells, text)
		}

		if len(cells) > 0 {
			rows = append(rows, cells)
		}
	}

	if len(rows) == 0 {
		return nil, fmt.Errorf("ячейки <td> не найдены в HTML таблице")
	}

	return parseRowsToXLSRows(rows), nil
}

func stripHTML(s string) string {
	// убираем все теги
	re := regexp.MustCompile(`(?is)<.*?>`)
	return re.ReplaceAllString(s, "")
}

//преобразуем строки в структуры XLSRow
func parseRowsToXLSRows(rows [][]string) []XLSRow {
	var xlsRows []XLSRow

	fmt.Printf("Найдено строк в таблице: %d\n", len(rows))

	for i, row := range rows {
		if i == 0 { // Пропускаем заголовки
			fmt.Printf("Заголовки: %v\n", row)
			continue
		}

		fmt.Printf("Строка %d: %v\n", i, row)

		// Заполняем структуру, проверяя границы массива
		xlsRow := XLSRow{}
		if len(row) > 0 {
			xlsRow.Surname = row[0]
		}
		if len(row) > 1 {
			xlsRow.Name = row[1]
		}
		if len(row) > 2 {
			xlsRow.Patronymic = row[2]
		}
		if len(row) > 3 {
			xlsRow.BirthYear = row[3]
		}
		if len(row) > 4 {
			xlsRow.BirthMonth = row[4]
		}
		if len(row) > 5 {
			xlsRow.BirthDay = row[5]
		}
		if len(row) > 6 {
			xlsRow.Result = row[6]
		}
		if len(row) > 7 {
			xlsRow.WantedPersons = row[7]
		}
		if len(row) > 8 {
			xlsRow.OSKRegion = row[8]
		}
		if len(row) > 9 {
			xlsRow.OSKGIAZ = row[9]
		}
		if len(row) > 10 {
			xlsRow.AdminPracticeR = row[10]
		}
		if len(row) > 11 {
			xlsRow.AdminPracticeF = row[11]
		}
		if len(row) > 12 {
			xlsRow.ZAGSDeath = row[12]
		}
		if len(row) > 13 {
			xlsRow.Restricted = row[13]
		}
		if len(row) > 14 {
			xlsRow.PassportRF = row[14]
		}
		if len(row) > 15 {
			xlsRow.DeportationMode = row[15]
		}

		xlsRows = append(xlsRows, xlsRow)
	}

	fmt.Printf("Преобразовано строк в XLSRow: %d\n", len(xlsRows))
	return xlsRows
}

func MatchXMLWithXLS(xmlData []byte, xlsRows []XLSRow) ([]XLSRow, error) {
	var list List
	err := xml.Unmarshal(xmlData, &list)
	if err != nil {
		return nil, err
	}

	xmlMap := make(map[string]string)

	for _, doc := range list.Document {
		p := doc.RequestInfo.ConvictionPerson
		xmlDate := strings.Split(p.CPBirthday, ".")
		if len(xmlDate) != 3 {
			continue
		}

		// Ключ для поиска: Фамилия_Имя_Отчество_Год_Месяц_День
		key1 := fmt.Sprintf("%s_%s_%s_%s_%s_%s",
			p.CPSurname, p.CPName, p.CPPatronymic,
			xmlDate[2], xmlDate[1], xmlDate[0])
		xmlMap[key1] = doc.DocNumber

		// Если есть старая фамилия, добавляем и ее
		if p.CPLastFIO != nil && p.CPLastFIO.CPLSurname != "" {
			key2 := fmt.Sprintf("%s_%s_%s_%s_%s_%s",
				p.CPLastFIO.CPLSurname, p.CPName, p.CPPatronymic,
				xmlDate[2], xmlDate[1], xmlDate[0])
			xmlMap[key2] = doc.DocNumber
		}
	}

	// Сравниваем с XLS данными
	for i, xlsRow := range xlsRows {
		key := fmt.Sprintf("%s_%s_%s_%s_%s_%s",
			xlsRow.Surname, xlsRow.Name, xlsRow.Patronymic,
			xlsRow.BirthYear, xlsRow.BirthMonth, xlsRow.BirthDay)

		if docNumber, exists := xmlMap[key]; exists {
			xlsRows[i].DocumentNumber = docNumber
		}
	}

	return xlsRows, nil
}

func ModifyXLSFile(filename string, xlsRows []XLSRow) error {
	return createNewExcelFile(filename, xlsRows)
}

func createNewExcelFile(filename string, xlsRows []XLSRow) error {
	f := excelize.NewFile()
	mainSheet := "Sheet1"
	positiveSheet := "Положительный результат"

	f.NewSheet(positiveSheet)

	// ===== Стили =====
	headerStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"D9E1F2"}},
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	gridStyle, _ := f.NewStyle(&excelize.Style{
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	highlightStyle, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FFFF00"}}, // желтый
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	// Стиль для жирной темной ячейки "ДА"
	darkDaCellStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FF6600"}}, // темно-оранжевый
		Border: []excelize.Border{
			{Type: "left", Color: "000000", Style: 1},
			{Type: "right", Color: "000000", Style: 1},
			{Type: "top", Color: "000000", Style: 1},
			{Type: "bottom", Color: "000000", Style: 1},
		},
	})

	headers := []string{
		"№ документа", "Фамилия", "Имя", "Отчество",
		"Год рождения", "Месяц рождения", "День рождения",
		"Результат", "Розыск лиц", "ОСК регион", "ОСК ГИАЦ",
		"Адмпрактика регион", "Адмпрактика ФИС-М",
		"ЗАГС рег.смерти", "Запретники", "Паспорт РФ", "Реж.высылки",
	}

	for i, header := range headers {
		cellMain, _ := excelize.CoordinatesToCellName(i+1, 1)
		cellPos, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(mainSheet, cellMain, header)
		f.SetCellValue(positiveSheet, cellPos, header)
		f.SetCellStyle(mainSheet, cellMain, cellMain, headerStyle)
		f.SetCellStyle(positiveSheet, cellPos, cellPos, headerStyle)
	}

	mainRowIndex := 2
	posRowIndex := 2

	colIndices := map[string]int{
		"Розыск лиц": 9,
		"ОСК регион": 10,
		"ОСК ГИАЦ":   11,
	}

	for _, row := range xlsRows {
		isPositive := row.WantedPersons == "ДА" || row.OSKRegion == "ДА" || row.OSKGIAZ == "ДА"

		rowData := []string{
			row.DocumentNumber,
			row.Surname,
			row.Name,
			row.Patronymic,
			row.BirthYear,
			row.BirthMonth,
			row.BirthDay,
			row.Result,
			row.WantedPersons,
			row.OSKRegion,
			row.OSKGIAZ,
			row.AdminPracticeR,
			row.AdminPracticeF,
			row.ZAGSDeath,
			row.Restricted,
			row.PassportRF,
			row.DeportationMode,
		}

		// --- Основной лист ---
		for j, value := range rowData {
			cell, _ := excelize.CoordinatesToCellName(j+1, mainRowIndex)
			f.SetCellValue(mainSheet, cell, value)
			f.SetCellStyle(mainSheet, cell, cell, gridStyle)

			// Если ячейка в столбцах с "ДА", делаем ее темной и жирной
			if (j+1 == colIndices["Розыск лиц"] && value == "ДА") ||
				(j+1 == colIndices["ОСК регион"] && value == "ДА") ||
				(j+1 == colIndices["ОСК ГИАЦ"] && value == "ДА") {
				f.SetCellStyle(mainSheet, cell, cell, darkDaCellStyle)
			}
		}

		// Жёлтая подсветка всей строки на основном листе
		startMain, _ := excelize.CoordinatesToCellName(1, mainRowIndex)
		endMain, _ := excelize.CoordinatesToCellName(len(headers), mainRowIndex)
		if isPositive {
			f.SetCellStyle(mainSheet, startMain, endMain, highlightStyle)
			for _, idx := range []int{colIndices["Розыск лиц"], colIndices["ОСК регион"], colIndices["ОСК ГИАЦ"]} {
				if idx >= 1 && idx <= len(rowData) && rowData[idx-1] == "ДА" {
					daCell, _ := excelize.CoordinatesToCellName(idx, mainRowIndex)
					f.SetCellStyle(mainSheet, daCell, daCell, darkDaCellStyle)
				}
			}
		}

		// --- Положительный результат ---
		if isPositive {
			for j, value := range rowData {
				cell, _ := excelize.CoordinatesToCellName(j+1, posRowIndex)
				f.SetCellValue(positiveSheet, cell, value)
				f.SetCellStyle(positiveSheet, cell, cell, gridStyle)
				if (j+1 == colIndices["Розыск лиц"] && value == "ДА") ||
					(j+1 == colIndices["ОСК регион"] && value == "ДА") ||
					(j+1 == colIndices["ОСК ГИАЦ"] && value == "ДА") {
					f.SetCellStyle(positiveSheet, cell, cell, darkDaCellStyle)
				}
			}
			posRowIndex++
		}

		mainRowIndex++
	}

	// Ширина колонок
	for i := 1; i <= len(headers); i++ {
		col, _ := excelize.ColumnNumberToName(i)
		f.SetColWidth(mainSheet, col, col, 15)
		f.SetColWidth(positiveSheet, col, col, 15)
	}

	now := time.Now()
	formattedTime := now.Format("02.01.2006_15-04-05")
	dir := filepath.Dir(filename)
	newFileName := filepath.Join(dir, "goususlugi_"+formattedTime+".xlsx")
	return f.SaveAs(newFileName)
}
