// Package analysis maps common Russian and English sales-column labels to
// revenue, cost and profit measures and calculates a transparent profitability
// summary for one Excel table.
package analysis

import (
	"errors"
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"unicode"

	"github.com/slchumak75-sketch/analitika-sales/internal/xlsx"
)

const (
	rankedItemsLimit = 5
	ModeItems        = "items"
	ModeReport       = "report"
)

type ColumnMapping struct {
	Role        string `json:"role"`
	Header      string `json:"header"`
	ExcelColumn string `json:"excelColumn"`
}

type LineItem struct {
	Label         string   `json:"label"`
	ExcelRow      int      `json:"excelRow"`
	Revenue       *float64 `json:"revenue,omitempty"`
	Cost          *float64 `json:"cost,omitempty"`
	Profit        float64  `json:"profit"`
	MarginPercent *float64 `json:"marginPercent,omitempty"`
}

type Report struct {
	Mode              string          `json:"mode"`
	SheetName         string          `json:"sheetName"`
	TableName         string          `json:"tableName"`
	Range             string          `json:"range"`
	StructuredTable   bool            `json:"structuredTable"`
	HeaderRow         int             `json:"headerRow"`
	HeaderExcelRow    int             `json:"headerExcelRow"`
	RowsInTable       int             `json:"rowsInTable"`
	RowsAnalyzed      int             `json:"rowsAnalyzed"`
	RowsSkipped       int             `json:"rowsSkipped"`
	TotalRowsSkipped  int             `json:"totalRowsSkipped"`
	RowsWithRevenue   int             `json:"rowsWithRevenue"`
	RowsWithCost      int             `json:"rowsWithCost"`
	RowsWithProfit    int             `json:"rowsWithProfit"`
	LossRows          int             `json:"lossRows"`
	ProfitableRows    int             `json:"profitableRows"`
	RevenueTotal      *float64        `json:"revenueTotal,omitempty"`
	CostTotal         *float64        `json:"costTotal,omitempty"`
	ProfitTotal       *float64        `json:"profitTotal,omitempty"`
	MarginPercent     *float64        `json:"marginPercent,omitempty"`
	GrossRevenue      *float64        `json:"grossRevenue,omitempty"`
	OutputVat         *float64        `json:"outputVat,omitempty"`
	VatPayable        *float64        `json:"vatPayable,omitempty"`
	IncomeTax         *float64        `json:"incomeTax,omitempty"`
	NetProfit         *float64        `json:"netProfit,omitempty"`
	NetMarginPercent  *float64        `json:"netMarginPercent,omitempty"`
	TotalColumnHeader string          `json:"totalColumnHeader,omitempty"`
	Columns           []ColumnMapping `json:"columns"`
	ReportLines       []ReportLine    `json:"reportLines,omitempty"`
	Periods           []ReportPeriod  `json:"periods,omitempty"`
	TopItems          []LineItem      `json:"topItems"`
	LowestItems       []LineItem      `json:"lowestItems"`
	Warnings          []string        `json:"warnings"`
}

type ReportLine struct {
	Label    string   `json:"label"`
	Role     string   `json:"role"`
	ExcelRow int      `json:"excelRow,omitempty"`
	Amount   *float64 `json:"amount,omitempty"`
	Subtotal bool     `json:"subtotal,omitempty"`
}

type ReportPeriod struct {
	Label            string   `json:"label"`
	ExcelColumn      string   `json:"excelColumn"`
	GrossRevenue     *float64 `json:"grossRevenue,omitempty"`
	Revenue          *float64 `json:"revenue,omitempty"`
	Cost             *float64 `json:"cost,omitempty"`
	Profit           *float64 `json:"profit,omitempty"`
	MarginPercent    *float64 `json:"marginPercent,omitempty"`
	IncomeTax        *float64 `json:"incomeTax,omitempty"`
	NetProfit        *float64 `json:"netProfit,omitempty"`
	NetMarginPercent *float64 `json:"netMarginPercent,omitempty"`
	VatPayable       *float64 `json:"vatPayable,omitempty"`
}

type columnKind int

const (
	kindNone columnKind = iota
	kindRevenue
	kindRevenueUnit
	kindCost
	kindCostUnit
	kindProfit
	kindProfitUnit
	kindProfitPercent
	kindQuantity
	kindProduct
)

type columnCandidate struct {
	index       int
	score       int
	header      string
	excelColumn string
}

type columnSelection struct {
	revenue       *columnCandidate
	revenueUnit   *columnCandidate
	cost          *columnCandidate
	costUnit      *columnCandidate
	profit        *columnCandidate
	profitUnit    *columnCandidate
	profitPercent *columnCandidate
	quantity      *columnCandidate
	product       *columnCandidate
}

type rowAmounts struct {
	revenue    float64
	cost       float64
	profit     float64
	hasRevenue bool
	hasCost    bool
	hasProfit  bool
}

type rankedRow struct {
	item  LineItem
	order int
}

// Analyze detects whether the extracted table is an item-level sales table or a
// multi-period profit-and-loss report, then calculates the matching report.
func Analyze(table xlsx.Table) (Report, error) {
	if report, ok := analyzeReport(table); ok {
		return report, nil
	}
	return analyzeItems(table)
}

// analyzeItems calculates a profitability report for a single item-level table.
func analyzeItems(table xlsx.Table) (Report, error) {
	if len(table.Rows) == 0 {
		return Report{}, errors.New("в выбранной таблице нет строк")
	}

	headerIndex, err := findHeaderRow(table)
	if err != nil {
		return Report{}, err
	}
	headers := table.Rows[headerIndex]
	selection := selectColumns(headers)
	if selection.revenue == nil && selection.revenueUnit == nil &&
		selection.cost == nil && selection.costUnit == nil &&
		selection.profit == nil && selection.profitUnit == nil &&
		selection.profitPercent == nil {
		return Report{}, errors.New("не удалось распознать финансовые столбцы. Проверьте заголовки: выручка/продажи, себестоимость/расходы или прибыль")
	}

	headerExcelRow := excelRowAt(table, headerIndex)
	report := Report{
		Mode:            ModeItems,
		SheetName:       table.SheetName,
		TableName:       table.TableName,
		Range:           table.Range,
		StructuredTable: table.Structured,
		HeaderRow:       headerIndex + 1,
		HeaderExcelRow:  headerExcelRow,
		RowsInTable:     len(table.Rows) - headerIndex - 1,
		TopItems:        []LineItem{},
		LowestItems:     []LineItem{},
		Columns:         makeColumnMappings(selection),
		Warnings:        []string{},
	}

	warnings := make(map[string]bool)
	addWarning := func(message string) { warnings[message] = true }
	if selection.quantity == nil && (selection.revenueUnit != nil || selection.costUnit != nil || selection.profitUnit != nil) {
		addWarning("В таблице есть цена или прибыль за единицу, но не найдено количество. Для таких строк принято количество 1.")
	}
	if selection.revenue == nil && selection.revenueUnit == nil {
		addWarning("Столбец выручки не распознан: прибыль можно оценить, но общая маржинальность недоступна.")
	}
	if selection.cost == nil && selection.costUnit == nil && selection.profit == nil && selection.profitUnit == nil && selection.profitPercent == nil {
		addWarning("Себестоимость и прибыль не распознаны. Для расчёта прибыли нужны выручка и себестоимость либо готовый столбец прибыли.")
	}

	var revenueTotal, costTotal, profitTotal float64
	var pairedRevenue, pairedProfit float64
	var ranked []rankedRow
	mismatchRows := 0

	for rowIndex := headerIndex + 1; rowIndex < len(table.Rows); rowIndex++ {
		row := table.Rows[rowIndex]
		if rowIsEmpty(row) || isRepeatedHeader(row, headers) {
			continue
		}
		if isTotalsRow(row, selection.product) {
			report.TotalRowsSkipped++
			continue
		}

		amounts, mismatch, unitQuantityAssumed := calculateRow(row, selection)
		if unitQuantityAssumed {
			addWarning("Для расчёта цены за единицу в строках без количества принято количество 1.")
		}
		if mismatch {
			mismatchRows++
		}
		if !amounts.hasProfit {
			report.RowsSkipped++
			continue
		}

		report.RowsAnalyzed++
		report.RowsWithProfit++
		profitTotal += amounts.profit
		if amounts.profit < 0 {
			report.LossRows++
		} else if amounts.profit > 0 {
			report.ProfitableRows++
		}
		if amounts.hasRevenue {
			report.RowsWithRevenue++
			revenueTotal += amounts.revenue
			pairedRevenue += amounts.revenue
			pairedProfit += amounts.profit
		}
		if amounts.hasCost {
			report.RowsWithCost++
			costTotal += amounts.cost
		}

		line := LineItem{
			Label:    rowLabel(row, selection.product, excelRowAt(table, rowIndex)),
			ExcelRow: excelRowAt(table, rowIndex),
			Profit:   amounts.profit,
		}
		if amounts.hasRevenue {
			line.Revenue = floatPointer(amounts.revenue)
		}
		if amounts.hasCost {
			line.Cost = floatPointer(amounts.cost)
		}
		if amounts.hasRevenue && amounts.revenue != 0 {
			line.MarginPercent = floatPointer(amounts.profit / amounts.revenue * 100)
		}
		ranked = append(ranked, rankedRow{item: line, order: rowIndex})
	}

	if report.RowsAnalyzed == 0 {
		return Report{}, errors.New("в таблице не удалось рассчитать прибыль ни по одной строке. Проверьте заголовки и числовые значения")
	}
	if report.RowsSkipped > 0 {
		addWarning(fmt.Sprintf("Не рассчитана прибыль для %d строк: в них не хватило числовых значений для выбранной схемы.", report.RowsSkipped))
	}
	if report.RowsWithRevenue < report.RowsWithProfit {
		addWarning("Выручка найдена не для всех строк с рассчитанной прибылью; общая маржинальность рассчитана только по строкам, где распознана выручка.")
	}
	if report.RowsWithCost < report.RowsWithProfit {
		addWarning("Себестоимость распознана не для всех строк с рассчитанной прибылью.")
	}
	if mismatchRows > 0 {
		addWarning(fmt.Sprintf("В %d строках значение столбца прибыли отличается от разницы выручки и себестоимости более чем на 5%% или 1 денежную единицу. Для отчёта использован готовый столбец прибыли.", mismatchRows))
	}

	if report.RowsWithRevenue > 0 {
		report.RevenueTotal = floatPointer(revenueTotal)
	}
	if report.RowsWithCost > 0 {
		report.CostTotal = floatPointer(costTotal)
	}
	report.ProfitTotal = floatPointer(profitTotal)
	if pairedRevenue != 0 {
		report.MarginPercent = floatPointer(pairedProfit / pairedRevenue * 100)
	} else if report.RowsWithRevenue > 0 {
		addWarning("Итоговая выручка равна нулю — процент маржинальности рассчитать нельзя.")
	}

	if report.ProfitTotal != nil && *report.ProfitTotal < 0 {
		addWarning("Суммарная прибыль отрицательная: по выбранным данным таблица показывает общий убыток.")
	}
	if report.LossRows > 0 {
		addWarning(fmt.Sprintf("Убыточными являются %d из %d рассчитанных строк.", report.LossRows, report.RowsAnalyzed))
	}
	if selection.profitPercent != nil && selection.profit == nil && selection.profitUnit == nil {
		addWarning("Прибыль оценена по столбцу процента маржи и выручке. Результат зависит от корректности процента в исходной таблице.")
	}

	sort.SliceStable(ranked, func(i, j int) bool {
		if ranked[i].item.Profit == ranked[j].item.Profit {
			return ranked[i].order < ranked[j].order
		}
		return ranked[i].item.Profit > ranked[j].item.Profit
	})
	for i := 0; i < len(ranked) && i < rankedItemsLimit; i++ {
		report.TopItems = append(report.TopItems, ranked[i].item)
	}
	sort.SliceStable(ranked, func(i, j int) bool {
		if ranked[i].item.Profit == ranked[j].item.Profit {
			return ranked[i].order < ranked[j].order
		}
		return ranked[i].item.Profit < ranked[j].item.Profit
	})
	for i := 0; i < len(ranked) && i < rankedItemsLimit; i++ {
		report.LowestItems = append(report.LowestItems, ranked[i].item)
	}

	for warning := range warnings {
		report.Warnings = append(report.Warnings, warning)
	}
	sort.Strings(report.Warnings)
	return report, nil
}

func findHeaderRow(table xlsx.Table) (int, error) {
	if table.HasHeader {
		if len(table.Rows) < 2 {
			return 0, errors.New("в таблице нет строк с данными под заголовком")
		}
		return 0, nil
	}
	limit := len(table.Rows)
	if limit > 25 {
		limit = 25
	}
	bestIndex, bestScore := -1, 0
	for rowIndex := 0; rowIndex < limit; rowIndex++ {
		score, financialCount := headerRowScore(table.Rows[rowIndex])
		if financialCount > 0 && score > bestScore {
			bestIndex, bestScore = rowIndex, score
		}
	}
	if bestIndex < 0 {
		return 0, errors.New("не найден заголовок с распознаваемым финансовым столбцом")
	}
	if len(table.Rows)-bestIndex < 2 {
		return 0, errors.New("под найденным заголовком нет строк с данными")
	}
	return bestIndex, nil
}

func headerRowScore(row []string) (score, financialCount int) {
	seen := make(map[columnKind]bool)
	for _, header := range row {
		kind, weight := classifyHeader(header)
		if kind == kindNone {
			continue
		}
		score += weight
		if kind == kindRevenue || kind == kindRevenueUnit || kind == kindCost || kind == kindCostUnit || kind == kindProfit || kind == kindProfitUnit || kind == kindProfitPercent {
			if !seen[kind] {
				financialCount++
				seen[kind] = true
			}
		}
	}
	return score, financialCount
}

func selectColumns(headers []string) columnSelection {
	var selection columnSelection
	for index, header := range headers {
		kind, score := classifyHeader(header)
		if kind == kindNone {
			continue
		}
		candidate := columnCandidate{
			index:       index,
			score:       score,
			header:      strings.TrimSpace(header),
			excelColumn: excelColumnName(index + 1),
		}
		var current **columnCandidate
		switch kind {
		case kindRevenue:
			current = &selection.revenue
		case kindRevenueUnit:
			current = &selection.revenueUnit
		case kindCost:
			current = &selection.cost
		case kindCostUnit:
			current = &selection.costUnit
		case kindProfit:
			current = &selection.profit
		case kindProfitUnit:
			current = &selection.profitUnit
		case kindProfitPercent:
			current = &selection.profitPercent
		case kindQuantity:
			current = &selection.quantity
		case kindProduct:
			current = &selection.product
		}
		if current != nil && (*current == nil || candidate.score > (*current).score) {
			copy := candidate
			*current = &copy
		}
	}
	return selection
}

func classifyHeader(header string) (columnKind, int) {
	normalized := normalizeHeader(header)
	if normalized == "" {
		return kindNone, 0
	}
	unit := containsAny(normalized, "за единицу", "за шт", "за штуку", "на единицу", "на ед", "на шт", "per unit", "unit price", "unit cost", "per item", "/шт", "/ед")
	percent := containsAny(normalized, "%", "процент", "рентабельност", "margin rate", "margin %", "маржинальност")

	if containsAny(normalized, "колич", "qty", "quantity", "штук", "шт.", "units sold", "count") {
		return kindQuantity, 4
	}
	if containsAny(normalized, "товар", "номенклатур", "наименован", "продукт", "артикул", "sku", "product", "item", "позиция") {
		return kindProduct, 3
	}
	if containsAny(normalized, "прибыл", "profit", "gross profit", "net income") {
		if percent {
			return kindProfitPercent, 7
		}
		if unit {
			return kindProfitUnit, 6
		}
		return kindProfit, 8
	}
	if containsAny(normalized, "маржинальност", "рентабельност", "gross margin", "margin rate") {
		return kindProfitPercent, 7
	}
	if containsAny(normalized, "маржа", "маржи", "margin") {
		if percent || containsAny(normalized, "margin %") {
			return kindProfitPercent, 6
		}
		return kindProfit, 5
	}
	if containsAny(normalized, "себестоим", "закуп", "cogs", "cost of goods", "purchase cost", "затрат", "расход", "cost", "издерж") {
		if unit || containsAny(normalized, "цена", "price", "за шт", "на ед") {
			return kindCostUnit, 7
		}
		return kindCost, 8
	}
	if containsAny(normalized, "выруч", "продаж", "реализац", "доход", "revenue", "sales", "turnover", "выручка", "реализовано") {
		if unit || containsAny(normalized, "цена", "price", "за шт", "на ед") {
			return kindRevenueUnit, 7
		}
		return kindRevenue, 8
	}
	if containsAny(normalized, "цена", "price") {
		return kindRevenueUnit, 2
	}
	if containsAny(normalized, "сумма", "amount", "total") && !containsAny(normalized, "итого по затрат", "total cost") {
		return kindRevenue, 2
	}
	return kindNone, 0
}

func calculateRow(row []string, selection columnSelection) (rowAmounts, bool, bool) {
	var amounts rowAmounts
	quantity, hasQuantity := readColumnNumber(row, selection.quantity)
	assumedQuantity := false
	if !hasQuantity {
		quantity = 1
		assumedQuantity = selection.revenueUnit != nil || selection.costUnit != nil || selection.profitUnit != nil
	}

	if value, ok := readColumnNumber(row, selection.revenue); ok {
		amounts.revenue, amounts.hasRevenue = value, true
	} else if value, ok := readColumnNumber(row, selection.revenueUnit); ok {
		amounts.revenue, amounts.hasRevenue = value*quantity, true
	}
	if value, ok := readColumnNumber(row, selection.cost); ok {
		amounts.cost, amounts.hasCost = value, true
	} else if value, ok := readColumnNumber(row, selection.costUnit); ok {
		amounts.cost, amounts.hasCost = value*quantity, true
	}

	if value, ok := readColumnNumber(row, selection.profit); ok {
		amounts.profit, amounts.hasProfit = value, true
	} else if value, ok := readColumnNumber(row, selection.profitUnit); ok {
		amounts.profit, amounts.hasProfit = value*quantity, true
	}

	var marginPercent float64
	hasMarginPercent := false
	if value, ok := readColumnNumber(row, selection.profitPercent); ok {
		marginPercent = value
		if !strings.Contains(rowValue(row, selection.profitPercent), "%") && math.Abs(marginPercent) <= 1 {
			// Excel stores percentage-formatted cells as fractions (for example,
			// 0.25 for 25%). A literal percent sign, in contrast, was parsed as
			// 25 and should not be scaled again.
			marginPercent *= 100
		}
		hasMarginPercent = true
	}

	mismatch := false
	if !amounts.hasProfit {
		if amounts.hasRevenue && amounts.hasCost {
			amounts.profit = amounts.revenue - amounts.cost
			amounts.hasProfit = true
		} else if amounts.hasRevenue && hasMarginPercent {
			amounts.profit = amounts.revenue * marginPercent / 100
			amounts.hasProfit = true
		}
	} else if amounts.hasRevenue && amounts.hasCost {
		calculated := amounts.revenue - amounts.cost
		tolerance := math.Max(1, math.Abs(calculated)*0.05)
		mismatch = math.Abs(amounts.profit-calculated) > tolerance
	}

	if amounts.hasProfit && amounts.hasRevenue && !amounts.hasCost {
		amounts.cost = amounts.revenue - amounts.profit
		amounts.hasCost = true
	}
	if amounts.hasProfit && amounts.hasCost && !amounts.hasRevenue {
		amounts.revenue = amounts.cost + amounts.profit
		amounts.hasRevenue = true
	}
	return amounts, mismatch, assumedQuantity
}

func makeColumnMappings(selection columnSelection) []ColumnMapping {
	var result []ColumnMapping
	appendMapping := func(role string, column *columnCandidate) {
		if column == nil {
			return
		}
		result = append(result, ColumnMapping{Role: role, Header: column.header, ExcelColumn: column.excelColumn})
	}
	appendMapping("Выручка", selection.revenue)
	appendMapping("Цена продажи за единицу", selection.revenueUnit)
	appendMapping("Себестоимость", selection.cost)
	appendMapping("Себестоимость за единицу", selection.costUnit)
	appendMapping("Прибыль", selection.profit)
	appendMapping("Прибыль за единицу", selection.profitUnit)
	if selection.profitPercent != nil {
		role := "Маржа, %"
		if selection.profit != nil || selection.profitUnit != nil {
			role += " (резерв)"
		}
		appendMapping(role, selection.profitPercent)
	}
	appendMapping("Количество", selection.quantity)
	appendMapping("Товар / позиция", selection.product)
	return result
}

func readColumnNumber(row []string, candidate *columnCandidate) (float64, bool) {
	if candidate == nil || candidate.index < 0 || candidate.index >= len(row) {
		return 0, false
	}
	return parseNumber(row[candidate.index])
}

func rowValue(row []string, candidate *columnCandidate) string {
	if candidate == nil || candidate.index < 0 || candidate.index >= len(row) {
		return ""
	}
	return row[candidate.index]
}

func rowIsEmpty(row []string) bool {
	for _, cell := range row {
		if strings.TrimSpace(cell) != "" {
			return false
		}
	}
	return true
}

func isRepeatedHeader(row, headers []string) bool {
	matched := 0
	checked := 0
	for index, header := range headers {
		if strings.TrimSpace(header) == "" || index >= len(row) {
			continue
		}
		checked++
		if normalizeHeader(header) == normalizeHeader(row[index]) {
			matched++
		}
	}
	return checked > 0 && matched == checked
}

func isTotalsRow(row []string, product *columnCandidate) bool {
	index := 0
	if product != nil {
		index = product.index
	}
	if index >= len(row) || strings.TrimSpace(row[index]) == "" {
		for candidateIndex, value := range row {
			if strings.TrimSpace(value) != "" {
				index = candidateIndex
				break
			}
		}
	}
	if index >= len(row) {
		return false
	}
	label := normalizeHeader(row[index])
	label = strings.Trim(label, ":;,.! ")
	return label == "итого" || label == "итог" || label == "всего" || label == "total" || label == "grand total" ||
		strings.HasPrefix(label, "итого по ") || strings.HasPrefix(label, "общий итог") || strings.HasPrefix(label, "итоговая сумма")
}

func rowLabel(row []string, product *columnCandidate, excelRow int) string {
	if product != nil && product.index >= 0 && product.index < len(row) {
		if label := strings.TrimSpace(row[product.index]); label != "" {
			return label
		}
	}
	return fmt.Sprintf("Строка %d", excelRow)
}

func excelRowAt(table xlsx.Table, index int) int {
	if index >= 0 && index < len(table.ExcelRows) && table.ExcelRows[index] > 0 {
		return table.ExcelRows[index]
	}
	return index + 1
}

func excelColumnName(column int) string {
	var result []byte
	for column > 0 {
		column--
		result = append(result, byte('A'+column%26))
		column /= 26
	}
	for left, right := 0, len(result)-1; left < right; left, right = left+1, right-1 {
		result[left], result[right] = result[right], result[left]
	}
	return string(result)
}

func normalizeHeader(value string) string {
	value = strings.ToLower(strings.TrimSpace(value))
	value = strings.ReplaceAll(value, "ё", "е")
	var builder strings.Builder
	space := false
	for _, character := range value {
		if unicode.IsSpace(character) {
			if !space {
				builder.WriteByte(' ')
				space = true
			}
			continue
		}
		builder.WriteRune(character)
		space = false
	}
	return strings.TrimSpace(builder.String())
}

func containsAny(value string, terms ...string) bool {
	for _, term := range terms {
		if strings.Contains(value, term) {
			return true
		}
	}
	return false
}

func parseNumber(input string) (float64, bool) {
	input = strings.TrimSpace(input)
	if input == "" || input == "-" || input == "—" || input == "–" {
		return 0, false
	}
	negative := strings.HasPrefix(input, "(") && strings.HasSuffix(input, ")")
	runes := []rune(input)
	var filtered strings.Builder
	for index, character := range runes {
		if character >= '0' && character <= '9' {
			filtered.WriteRune(character)
			continue
		}
		if (character == '.' || character == ',') && nearestDigit(runes, index, -1) && nearestDigit(runes, index, 1) {
			filtered.WriteRune(character)
			continue
		}
		if (character == '-' || character == '+') && filtered.Len() == 0 {
			filtered.WriteRune(character)
		}
	}
	value := filtered.String()
	if value == "" {
		return 0, false
	}
	if negative {
		value = strings.Trim(value, "()+-")
		value = "-" + value
	}

	commaCount := strings.Count(value, ",")
	dotCount := strings.Count(value, ".")
	switch {
	case commaCount > 0 && dotCount > 0:
		if strings.LastIndex(value, ",") > strings.LastIndex(value, ".") {
			value = strings.ReplaceAll(value, ".", "")
			value = strings.ReplaceAll(value, ",", ".")
		} else {
			value = strings.ReplaceAll(value, ",", "")
		}
	case commaCount > 0:
		value = normalizeRepeatedSeparator(value, ',')
	case dotCount > 1:
		value = normalizeRepeatedSeparator(value, '.')
	}
	parsed, err := strconv.ParseFloat(value, 64)
	if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) {
		return 0, false
	}
	return parsed, true
}

func nearestDigit(runes []rune, index, direction int) bool {
	for cursor := index + direction; cursor >= 0 && cursor < len(runes); cursor += direction {
		if unicode.IsSpace(runes[cursor]) {
			continue
		}
		return runes[cursor] >= '0' && runes[cursor] <= '9'
	}
	return false
}

func normalizeRepeatedSeparator(value string, separator rune) string {
	count := strings.Count(value, string(separator))
	if count == 1 {
		return strings.Replace(value, string(separator), ".", 1)
	}
	last := strings.LastIndex(value, string(separator))
	fractionDigits := len(value) - last - 1
	// Repeated groups of three digits are treated as thousands separators;
	// otherwise the last separator is used as the decimal mark.
	if fractionDigits == 3 && allGroupsHaveThreeDigits(value, separator) {
		return strings.ReplaceAll(value, string(separator), "")
	}
	integer := strings.ReplaceAll(value[:last], string(separator), "")
	fraction := value[last+1:]
	return integer + "." + fraction
}

func allGroupsHaveThreeDigits(value string, separator rune) bool {
	parts := strings.Split(value, string(separator))
	if len(parts) < 2 || len(parts[0]) == 0 || len(parts[0]) > 3 {
		return false
	}
	for _, part := range parts[1:] {
		if len(part) != 3 {
			return false
		}
	}
	return true
}

func floatPointer(value float64) *float64 {
	copy := value
	return &copy
}
