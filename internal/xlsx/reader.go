// Package xlsx reads the first structured Excel table in a workbook, or (if the
// workbook has no structured tables) the used range of its first non-empty sheet.
// It intentionally uses only the Go standard library so Windows builds are
// reproducible without native dependencies.
package xlsx

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"strconv"
	"strings"
)

const (
	maxExpandedWorkbookSize = 256 << 20
	maxZipEntrySize         = 128 << 20
	maxTableRows            = 100_000
	maxTableColumns         = 512
	maxWorksheetCells       = 1_000_000
	maxTableCells           = 1_000_000
)

type Table struct {
	SheetName  string
	TableName  string
	Range      string
	Rows       [][]string
	ExcelRows  []int
	HasHeader  bool
	Structured bool
}

type workbookSheet struct {
	name  string
	state string
	path  string
}

type parsedSheet struct {
	meta        workbookSheet
	cells       map[int]map[int]string
	tableRelIDs []string
}

type rawCell struct {
	Ref    string `xml:"r,attr"`
	Type   string `xml:"t,attr"`
	Value  string `xml:"v"`
	Inline struct {
		Text string `xml:"t"`
		Runs []struct {
			Text string `xml:"t"`
		} `xml:"r"`
	} `xml:"is"`
}

// ReadFirstTable parses an XLSX/XLSM byte slice. Formula cells use their cached
// value; this reader does not recalculate Excel formulas.
func ReadFirstTable(data []byte) (Table, error) {
	if len(data) == 0 {
		return Table{}, errors.New("файл пуст")
	}

	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return Table{}, errors.New("файл не является корректной книгой XLSX/XLSM")
	}

	files := make(map[string]*zip.File, len(archive.File))
	var expandedSize uint64
	for _, file := range archive.File {
		name := path.Clean(strings.TrimPrefix(file.Name, "/"))
		if name == "." || strings.HasPrefix(name, "../") {
			continue
		}
		if file.UncompressedSize64 > maxZipEntrySize {
			return Table{}, errors.New("в книге есть слишком большой элемент; размер распаковки ограничен")
		}
		expandedSize += file.UncompressedSize64
		if expandedSize > maxExpandedWorkbookSize {
			return Table{}, errors.New("размер распакованной книги превышает безопасный лимит 256 МиБ")
		}
		files[name] = file
	}

	sharedStrings, err := readSharedStrings(files)
	if err != nil {
		return Table{}, fmt.Errorf("не удалось прочитать текстовые значения книги: %w", err)
	}

	sheets, err := readWorkbookSheets(files)
	if err != nil {
		return Table{}, err
	}

	var fallback *parsedSheet
	for _, sheet := range sheets {
		if sheet.state == "hidden" || sheet.state == "veryHidden" {
			continue
		}
		item, err := parseWorksheet(sheet, files, sharedStrings)
		if err != nil {
			// A damaged or unsupported sheet should not prevent us from using the
			// next visible worksheet that contains a valid table.
			continue
		}
		if table, found := firstStructuredTable(item, files); found {
			return table, nil
		}
		if fallback == nil && len(item.cells) > 0 {
			copy := item
			fallback = &copy
		}
	}

	// If the workbook contains no formal Excel table objects, use the first
	// visible non-empty worksheet as the first data table.
	if fallback != nil {
		if table, found, err := sheetUsedRange(*fallback); err != nil {
			return Table{}, err
		} else if found {
			return table, nil
		}
	}

	return Table{}, errors.New("в книге не найдена непустая таблица или лист с данными")
}

func readZipFile(files map[string]*zip.File, name string) ([]byte, error) {
	file := files[path.Clean(name)]
	if file == nil {
		return nil, fmt.Errorf("элемент %q отсутствует", name)
	}
	if file.UncompressedSize64 > maxZipEntrySize {
		return nil, fmt.Errorf("элемент %q превышает безопасный лимит", name)
	}
	reader, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer reader.Close()
	data, err := io.ReadAll(io.LimitReader(reader, maxZipEntrySize+1))
	if err != nil {
		return nil, err
	}
	if len(data) > maxZipEntrySize {
		return nil, fmt.Errorf("элемент %q превышает безопасный лимит", name)
	}
	return data, nil
}

func readSharedStrings(files map[string]*zip.File) ([]string, error) {
	file := files["xl/sharedStrings.xml"]
	if file == nil {
		return nil, nil
	}
	data, err := readZipFile(files, "xl/sharedStrings.xml")
	if err != nil {
		return nil, err
	}
	var doc struct {
		Items []struct {
			Text string `xml:"t"`
			Runs []struct {
				Text string `xml:"t"`
			} `xml:"r"`
		} `xml:"si"`
	}
	if err := xml.Unmarshal(data, &doc); err != nil {
		return nil, err
	}
	result := make([]string, 0, len(doc.Items))
	for _, item := range doc.Items {
		text := item.Text
		if len(item.Runs) > 0 {
			var builder strings.Builder
			for _, run := range item.Runs {
				builder.WriteString(run.Text)
			}
			text = builder.String()
		}
		result = append(result, text)
	}
	return result, nil
}

func readWorkbookSheets(files map[string]*zip.File) ([]workbookSheet, error) {
	workbookData, err := readZipFile(files, "xl/workbook.xml")
	if err != nil {
		return nil, errors.New("в XLSX не найден файл xl/workbook.xml")
	}

	relationships := map[string]string{}
	if relationshipData, readErr := readZipFile(files, "xl/_rels/workbook.xml.rels"); readErr == nil {
		relationships = parseRelationships(relationshipData)
	}

	decoder := xml.NewDecoder(bytes.NewReader(workbookData))
	var sheets []workbookSheet
	for {
		token, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		}
		if tokenErr != nil {
			return nil, fmt.Errorf("не удалось разобрать список листов: %w", tokenErr)
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "sheet" {
			continue
		}
		name := attribute(start, "name")
		relID := attribute(start, "id")
		state := attribute(start, "state")
		sheetPath := ""
		if target := relationships[relID]; target != "" {
			sheetPath = resolveTarget("xl/workbook.xml", target)
		}
		if sheetPath == "" {
			// A few non-conforming workbooks omit relationship metadata. The
			// standard sheet numbering is a useful best-effort fallback.
			sheetPath = fmt.Sprintf("xl/worksheets/sheet%d.xml", len(sheets)+1)
		}
		sheets = append(sheets, workbookSheet{name: name, state: state, path: sheetPath})
	}
	if len(sheets) == 0 {
		return nil, errors.New("в книге не найдено ни одного листа")
	}
	return sheets, nil
}

func parseWorksheet(meta workbookSheet, files map[string]*zip.File, sharedStrings []string) (parsedSheet, error) {
	data, err := readZipFile(files, meta.path)
	if err != nil {
		return parsedSheet{}, err
	}
	result := parsedSheet{meta: meta, cells: make(map[int]map[int]string)}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	currentRow := 0
	fallbackRow := 0
	cellCount := 0
	for {
		token, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		}
		if tokenErr != nil {
			return parsedSheet{}, tokenErr
		}
		switch element := token.(type) {
		case xml.StartElement:
			switch element.Name.Local {
			case "row":
				fallbackRow++
				currentRow = intAttribute(element, "r")
				if currentRow == 0 {
					currentRow = fallbackRow
				}
			case "c":
				var cell rawCell
				if err := decoder.DecodeElement(&cell, &element); err != nil {
					return parsedSheet{}, err
				}
				row, column, ok := parseCellReference(cell.Ref)
				if !ok {
					row, column = currentRow, 0
				}
				if row <= 0 {
					continue
				}
				value := cellValue(cell, sharedStrings)
				if strings.TrimSpace(value) == "" {
					continue
				}
				cellCount++
				if cellCount > maxWorksheetCells {
					return parsedSheet{}, errors.New("лист превышает лимит 1 000 000 непустых ячеек")
				}
				if result.cells[row] == nil {
					result.cells[row] = make(map[int]string)
				}
				result.cells[row][column] = value
			case "tablePart":
				if relID := attribute(element, "id"); relID != "" {
					result.tableRelIDs = append(result.tableRelIDs, relID)
				}
			}
		}
	}
	return result, nil
}

func cellValue(cell rawCell, sharedStrings []string) string {
	switch cell.Type {
	case "s":
		index, err := strconv.Atoi(strings.TrimSpace(cell.Value))
		if err == nil && index >= 0 && index < len(sharedStrings) {
			return sharedStrings[index]
		}
		return ""
	case "inlineStr":
		if len(cell.Inline.Runs) == 0 {
			return cell.Inline.Text
		}
		var builder strings.Builder
		for _, run := range cell.Inline.Runs {
			builder.WriteString(run.Text)
		}
		return builder.String()
	case "b":
		if cell.Value == "1" {
			return "TRUE"
		}
		if cell.Value == "0" {
			return "FALSE"
		}
	}
	return strings.TrimSpace(cell.Value)
}

func firstStructuredTable(sheet parsedSheet, files map[string]*zip.File) (Table, bool) {
	if len(sheet.tableRelIDs) == 0 {
		return Table{}, false
	}
	relPath := relationshipPath(sheet.meta.path)
	relData, err := readZipFile(files, relPath)
	if err != nil {
		return Table{}, false
	}
	relationships := parseRelationships(relData)
	for _, relID := range sheet.tableRelIDs {
		target := relationships[relID]
		if target == "" {
			continue
		}
		tablePath := resolveTarget(sheet.meta.path, target)
		tableData, readErr := readZipFile(files, tablePath)
		if readErr != nil {
			continue
		}
		ref, name, hasHeader, ok := parseTableMetadata(tableData)
		if !ok {
			continue
		}
		table, buildErr := buildRange(sheet, ref, name, true, hasHeader)
		if buildErr == nil {
			return table, true
		}
	}
	return Table{}, false
}

func parseTableMetadata(data []byte) (ref, name string, hasHeader, ok bool) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return "", "", false, false
		}
		if err != nil {
			return "", "", false, false
		}
		start, found := token.(xml.StartElement)
		if !found || start.Name.Local != "table" {
			continue
		}
		ref = attribute(start, "ref")
		name = attribute(start, "displayName")
		if name == "" {
			name = attribute(start, "name")
		}
		headerCount := intAttribute(start, "headerRowCount")
		hasHeader = true // OOXML defaults headerRowCount to 1.
		for _, attr := range start.Attr {
			if attr.Name.Local == "headerRowCount" {
				hasHeader = headerCount > 0
				break
			}
		}
		return ref, name, hasHeader, ref != ""
	}
}

func sheetUsedRange(sheet parsedSheet) (Table, bool, error) {
	minRow, minColumn, maxRow, maxColumn, found := cellBounds(sheet.cells)
	if !found {
		return Table{}, false, nil
	}
	ref := cellReference(minColumn, minRow) + ":" + cellReference(maxColumn, maxRow)
	table, err := buildRange(sheet, ref, "", false, false)
	if err != nil {
		return Table{}, false, err
	}
	return table, true, nil
}

func buildRange(sheet parsedSheet, ref, name string, structured, hasHeader bool) (Table, error) {
	minColumn, minRow, maxColumn, maxRow, ok := parseRange(ref)
	if !ok {
		return Table{}, fmt.Errorf("не удалось разобрать диапазон таблицы %q", ref)
	}
	width := maxColumn - minColumn + 1
	height := maxRow - minRow + 1
	if width <= 0 || height <= 0 || width > maxTableColumns || height > maxTableRows || width > maxTableCells/height {
		return Table{}, fmt.Errorf("размер таблицы %q превышает безопасный лимит", ref)
	}
	rows := make([][]string, height)
	excelRows := make([]int, height)
	for rowOffset := 0; rowOffset < height; rowOffset++ {
		excelRow := minRow + rowOffset
		excelRows[rowOffset] = excelRow
		rows[rowOffset] = make([]string, width)
		for columnOffset := 0; columnOffset < width; columnOffset++ {
			if value := sheet.cells[excelRow][minColumn+columnOffset]; value != "" {
				rows[rowOffset][columnOffset] = value
			}
		}
	}
	if name == "" {
		name = "Первый диапазон данных"
	}
	return Table{
		SheetName:  sheet.meta.name,
		TableName:  name,
		Range:      cellReference(minColumn, minRow) + ":" + cellReference(maxColumn, maxRow),
		Rows:       rows,
		ExcelRows:  excelRows,
		HasHeader:  hasHeader,
		Structured: structured,
	}, nil
}

func cellBounds(cells map[int]map[int]string) (minRow, minColumn, maxRow, maxColumn int, found bool) {
	for row, columns := range cells {
		for column, value := range columns {
			if strings.TrimSpace(value) == "" || column <= 0 {
				continue
			}
			if !found {
				minRow, maxRow = row, row
				minColumn, maxColumn = column, column
				found = true
				continue
			}
			if row < minRow {
				minRow = row
			}
			if row > maxRow {
				maxRow = row
			}
			if column < minColumn {
				minColumn = column
			}
			if column > maxColumn {
				maxColumn = column
			}
		}
	}
	return
}

func parseRelationships(data []byte) map[string]string {
	result := make(map[string]string)
	decoder := xml.NewDecoder(bytes.NewReader(data))
	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		start, ok := token.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		id := attribute(start, "Id")
		target := attribute(start, "Target")
		if id != "" && target != "" {
			result[id] = target
		}
	}
	return result
}

func parseCellReference(ref string) (row, column int, ok bool) {
	ref = strings.TrimSpace(ref)
	if bang := strings.LastIndex(ref, "!"); bang >= 0 {
		ref = ref[bang+1:]
	}
	ref = strings.ReplaceAll(ref, "$", "")
	letters := 0
	for letters < len(ref) {
		b := ref[letters]
		if b >= 'a' && b <= 'z' {
			b -= 'a' - 'A'
		}
		if b < 'A' || b > 'Z' {
			break
		}
		column = column*26 + int(b-'A'+1)
		letters++
	}
	if letters == 0 || letters == len(ref) {
		return 0, 0, false
	}
	parsedRow, err := strconv.Atoi(ref[letters:])
	if err != nil || parsedRow <= 0 || column <= 0 {
		return 0, 0, false
	}
	return parsedRow, column, true
}

func parseRange(ref string) (minColumn, minRow, maxColumn, maxRow int, ok bool) {
	ref = strings.TrimSpace(ref)
	if bang := strings.LastIndex(ref, "!"); bang >= 0 {
		ref = ref[bang+1:]
	}
	parts := strings.Split(ref, ":")
	if len(parts) == 1 {
		parts = append(parts, parts[0])
	}
	if len(parts) != 2 {
		return 0, 0, 0, 0, false
	}
	row1, column1, valid1 := parseCellReference(parts[0])
	row2, column2, valid2 := parseCellReference(parts[1])
	if !valid1 || !valid2 {
		return 0, 0, 0, 0, false
	}
	minColumn, maxColumn = column1, column2
	minRow, maxRow = row1, row2
	if minColumn > maxColumn {
		minColumn, maxColumn = maxColumn, minColumn
	}
	if minRow > maxRow {
		minRow, maxRow = maxRow, minRow
	}
	return minColumn, minRow, maxColumn, maxRow, true
}

func cellReference(column, row int) string {
	var letters []byte
	for column > 0 {
		column--
		letters = append(letters, byte('A'+column%26))
		column /= 26
	}
	for left, right := 0, len(letters)-1; left < right; left, right = left+1, right-1 {
		letters[left], letters[right] = letters[right], letters[left]
	}
	return string(letters) + strconv.Itoa(row)
}

func resolveTarget(sourcePath, target string) string {
	if strings.HasPrefix(target, "/") {
		return path.Clean(strings.TrimPrefix(target, "/"))
	}
	resolved := path.Clean(path.Join(path.Dir(sourcePath), target))
	if resolved == ".." || strings.HasPrefix(resolved, "../") {
		return ""
	}
	return resolved
}

func relationshipPath(sourcePath string) string {
	return path.Join(path.Dir(sourcePath), "_rels", path.Base(sourcePath)+".rels")
}

func attribute(element xml.StartElement, localName string) string {
	for _, attr := range element.Attr {
		if attr.Name.Local == localName {
			return attr.Value
		}
	}
	return ""
}

func intAttribute(element xml.StartElement, localName string) int {
	value := attribute(element, localName)
	parsed, _ := strconv.Atoi(value)
	return parsed
}
