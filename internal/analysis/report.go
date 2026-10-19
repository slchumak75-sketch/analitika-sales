package analysis

import (
	"fmt"
	"math"
	"sort"
	"strconv"
	"strings"
	"unicode"

	"github.com/slchumak75-sketch/analitika-sales/internal/xlsx"
)

// reportRole describes a recognized line of a profit-and-loss report.
type reportRole int

const (
	reportRoleNone reportRole = iota
	reportRoleRevenueTotal
	reportRoleRevenueComponent
	reportRoleCost
	reportRoleCostTotal
	reportRoleExpense
	reportRoleExpenseTotal
	reportRoleOutputVatTotal
	reportRoleOutputVatComponent
	reportRoleVatInput
	reportRoleVatPayable
	reportRoleIncomeTax
	reportRoleNetProfit
	reportRoleProfitBeforeTax
)

func (role reportRole) isRevenue() bool {
	return role == reportRoleRevenueTotal || role == reportRoleRevenueComponent
}

func (role reportRole) isCost() bool {
	return role == reportRoleCost || role == reportRoleCostTotal
}

func (role reportRole) isExpense() bool {
	return role == reportRoleExpense || role == reportRoleExpenseTotal
}

type reportSourceLine struct {
	role     reportRole
	label    string
	excelRow int
	values   map[int]float64
}

type reportLayout struct {
	headerIndex  int
	startColumn  int
	labelColumn  int
	totalColumn  int
	valueColumns []int
	lines        []reportSourceLine
}

func (layout reportLayout) excelColumn(relative int) string {
	start := layout.startColumn
	if start == 0 {
		start = 1
	}
	return excelColumnName(start + relative)
}

func (layout reportLayout) periodColumns() []int {
	periods := make([]int, 0, len(layout.valueColumns))
	for _, column := range layout.valueColumns {
		if column == layout.totalColumn {
			continue
		}
		periods = append(periods, column)
	}
	return periods
}

// analyzeReport recognizes a vertical profit-and-loss report with line labels in
// one column and periods in adjacent columns. It returns false when the table is
// not structured like such a report, so the item-level analyzer can handle it.
func analyzeReport(table xlsx.Table) (Report, bool) {
	layout, ok := detectReportLayout(table)
	if !ok {
		return Report{}, false
	}
	return buildReport(table, layout), true
}

func detectReportLayout(table xlsx.Table) (reportLayout, bool) {
	if len(table.Rows) < 4 {
		return reportLayout{}, false
	}
	headerLimit := len(table.Rows) - 1
	if headerLimit > 30 {
		headerLimit = 30
	}

	best := reportLayout{totalColumn: -1}
	bestScore := 0
	for headerIndex := 0; headerIndex < headerLimit; headerIndex++ {
		headerRow := table.Rows[headerIndex]
		if len(headerRow) < 3 {
			continue
		}
		totalColumn := reportTotalHeaderIndex(headerRow)
		labelColumn, roleCount := reportLabelColumn(table, headerIndex)
		if labelColumn < 0 || roleCount < 3 {
			continue
		}
		valueColumns := reportValueColumns(table, headerIndex, labelColumn)
		if len(valueColumns) < 2 {
			continue
		}
		if totalColumn >= 0 && !containsReportColumn(valueColumns, totalColumn) {
			totalColumn = -1
		}
		hasRevenue, hasCostSource, recognized := reportLayoutSignals(table, headerIndex, labelColumn)
		if !hasRevenue || !hasCostSource || recognized < 3 {
			continue
		}
		score := recognized*10 + len(valueColumns)
		if totalColumn >= 0 {
			score += 5
		}
		if score > bestScore {
			bestScore = score
			best = reportLayout{
				headerIndex:  headerIndex,
				startColumn:  tableStartColumn(table.Range),
				labelColumn:  labelColumn,
				totalColumn:  totalColumn,
				valueColumns: valueColumns,
			}
		}
	}
	if bestScore == 0 {
		return reportLayout{}, false
	}
	best.lines = reportSourceLines(table, best)
	if len(best.lines) < 3 {
		return reportLayout{}, false
	}
	return best, true
}

func reportTotalHeaderIndex(row []string) int {
	for column, cell := range row {
		normalized := normalizeHeader(cell)
		if normalized == "" || !reportHeaderLooksLikeLabel(cell) {
			continue
		}
		if containsAny(normalized, "итого", "всего", "total", "grand total") {
			return column
		}
	}
	return -1
}

func reportLabelColumn(table xlsx.Table, headerIndex int) (int, int) {
	if headerIndex+1 >= len(table.Rows) {
		return -1, 0
	}
	width := 0
	for _, row := range table.Rows {
		if len(row) > width {
			width = len(row)
		}
	}
	bestColumn, bestRoles, bestScore := -1, 0, 0
	for column := 0; column < width; column++ {
		roleCount, textCount := 0, 0
		for rowIndex := headerIndex + 1; rowIndex < len(table.Rows); rowIndex++ {
			cell := strings.TrimSpace(reportCell(table.Rows[rowIndex], column))
			if cell == "" {
				continue
			}
			textCount++
			if classifyReportLabel(cell) != reportRoleNone {
				roleCount++
			}
		}
		if roleCount == 0 {
			continue
		}
		score := roleCount*1000 + textCount
		if score > bestScore {
			bestColumn, bestRoles, bestScore = column, roleCount, score
		}
	}
	return bestColumn, bestRoles
}

func reportValueColumns(table xlsx.Table, headerIndex, labelColumn int) []int {
	headerRow := table.Rows[headerIndex]
	var columns []int
	for column := 0; column < len(headerRow); column++ {
		if column == labelColumn {
			continue
		}
		header := strings.TrimSpace(reportCell(headerRow, column))
		if header == "" || !reportHeaderLooksLikeLabel(header) {
			continue
		}
		numericCells := 0
		for rowIndex := headerIndex + 1; rowIndex < len(table.Rows); rowIndex++ {
			if _, ok := parseNumber(reportCell(table.Rows[rowIndex], column)); ok {
				numericCells++
				if numericCells >= 2 {
					break
				}
			}
		}
		if numericCells >= 2 {
			columns = append(columns, column)
		}
	}
	return columns
}

func reportHeaderLooksLikeLabel(value string) bool {
	value = strings.TrimSpace(value)
	if value == "" {
		return false
	}
	for _, character := range value {
		if unicode.IsLetter(character) {
			return true
		}
	}
	// A bare year is a common period header; arbitrary numeric cells from a
	// data row are not accepted as headers.
	if len(value) == 4 {
		year, err := strconv.Atoi(value)
		return err == nil && year >= 1900 && year <= 2100
	}
	return false
}

func reportLayoutSignals(table xlsx.Table, headerIndex, labelColumn int) (hasRevenue, hasCostSource bool, recognized int) {
	for rowIndex := headerIndex + 1; rowIndex < len(table.Rows); rowIndex++ {
		role := classifyReportLabel(reportCell(table.Rows[rowIndex], labelColumn))
		if role == reportRoleNone {
			continue
		}
		recognized++
		if role.isRevenue() {
			hasRevenue = true
		}
		if role.isCost() || role.isExpense() {
			hasCostSource = true
		}
	}
	return hasRevenue, hasCostSource, recognized
}

func reportSourceLines(table xlsx.Table, layout reportLayout) []reportSourceLine {
	var lines []reportSourceLine
	for rowIndex := layout.headerIndex + 1; rowIndex < len(table.Rows); rowIndex++ {
		label := strings.TrimSpace(reportCell(table.Rows[rowIndex], layout.labelColumn))
		role := classifyReportLabel(label)
		if role == reportRoleNone {
			continue
		}
		values := make(map[int]float64)
		for _, column := range layout.valueColumns {
			if value, ok := parseNumber(reportCell(table.Rows[rowIndex], column)); ok {
				values[column] = value
			}
		}
		if len(values) == 0 {
			continue
		}
		lines = append(lines, reportSourceLine{
			role:     role,
			label:    label,
			excelRow: excelRowAt(table, rowIndex),
			values:   values,
		})
	}
	return lines
}

func classifyReportLabel(label string) reportRole {
	normalized := normalizeHeader(label)
	if normalized == "" {
		return reportRoleNone
	}
	if containsAny(normalized, "без ндс") {
		// Such lines are usually components or deductions already represented by
		// the gross revenue and VAT lines; using them again would double count.
		return reportRoleNone
	}
	if containsAny(normalized, "net profit", "net income") || (containsAny(normalized, "чист") && containsAny(normalized, "прибыл", "доход")) {
		return reportRoleNetProfit
	}
	if containsAny(normalized, "налог на прибыл", "налог с прибыл", "income tax", "profit tax", "tax on profit") {
		return reportRoleIncomeTax
	}
	if containsAny(normalized, "ндс", "vat") {
		if containsAny(normalized, "к уплате", "payable", "налог ндс", "vat payable") {
			return reportRoleVatPayable
		}
		if containsAny(normalized, "входящ", "input", "к вычету") {
			return reportRoleVatInput
		}
		if containsAny(normalized, "вычет") {
			return reportRoleNone
		}
		if containsAny(normalized, "итого", "всего", "total") {
			return reportRoleOutputVatTotal
		}
		return reportRoleOutputVatComponent
	}
	if containsAny(normalized, "себестоим", "cogs", "cost of goods", "cost of sales") {
		if containsAny(normalized, "итого", "всего", "total") {
			return reportRoleCostTotal
		}
		if containsAny(normalized, "затрат", "расход") {
			return reportRoleExpense
		}
		return reportRoleCost
	}
	if containsAny(normalized, "выручк", "доход", "revenue", "sales", "turnover", "продаж", "реализац") {
		if containsAny(normalized, "товар", "вычет") {
			return reportRoleNone
		}
		if containsAny(normalized, "итого", "всего", "total") {
			return reportRoleRevenueTotal
		}
		return reportRoleRevenueComponent
	}
	if containsAny(normalized, "итого", "всего", "total") && containsAny(normalized, "затрат", "расход", "издерж", "expense", "opex") {
		return reportRoleExpenseTotal
	}
	if containsAny(normalized, "затрат", "расход", "издерж", "expense", "opex") {
		return reportRoleExpense
	}
	if containsAny(normalized, "прибыл", "profit") {
		if containsAny(normalized, "валов", "gross profit") {
			return reportRoleNone
		}
		return reportRoleProfitBeforeTax
	}
	return reportRoleNone
}

func buildReport(table xlsx.Table, layout reportLayout) Report {
	revenueLines := selectReportLines(layout.lines, reportRoleRevenueTotal, reportRoleRevenueComponent)
	costLines := selectReportLines(layout.lines, reportRoleCostTotal, reportRoleCost)
	expenseLines := selectReportLines(layout.lines, reportRoleExpenseTotal, reportRoleExpense)
	outputVatLines := selectReportLines(layout.lines, reportRoleOutputVatTotal, reportRoleOutputVatComponent)
	vatPayableLines := reportLinesWithRole(layout.lines, reportRoleVatPayable)
	incomeTaxLines := reportLinesWithRole(layout.lines, reportRoleIncomeTax)
	netProfitLines := reportLinesWithRole(layout.lines, reportRoleNetProfit)
	profitBeforeTaxLines := reportLinesWithRole(layout.lines, reportRoleProfitBeforeTax)

	warnings := make(map[string]bool)
	addWarning := func(message string) { warnings[message] = true }

	grossRevenue, hasRevenue := sumReportLines(revenueLines, layout, -1)
	outputVat, hasOutputVat := sumReportLines(outputVatLines, layout, -1)
	netRevenue := grossRevenue
	if hasOutputVat {
		netRevenue -= outputVat
	}
	cost, hasCost := sumReportLines(costLines, layout, -1)
	expense, hasExpense := sumReportLines(expenseLines, layout, -1)
	totalCost := cost + expense
	profit := netRevenue - totalCost
	vatPayable, hasVatPayable := sumReportLines(vatPayableLines, layout, -1)
	incomeTax, hasIncomeTax := sumReportLines(incomeTaxLines, layout, -1)

	netProfit := profit
	hasNetProfit := false
	if len(netProfitLines) > 0 {
		if value, ok := sumReportLines(netProfitLines, layout, -1); ok {
			netProfit = value
			hasNetProfit = true
		}
	}
	if !hasNetProfit && hasIncomeTax {
		netProfit = profit - incomeTax
		hasNetProfit = true
	}

	if len(profitBeforeTaxLines) > 0 {
		if stated, ok := sumReportLines(profitBeforeTaxLines, layout, -1); ok {
			tolerance := math.Max(1, math.Abs(profit)*0.01)
			if math.Abs(stated-profit) > tolerance {
				addWarning(fmt.Sprintf("Строка «%s» показывает прибыль %s, а расчёт по выручке и затратам — %s. В отчёте использован расчёт по выручке и затратам.", profitBeforeTaxLines[0].label, formatReportAmount(stated), formatReportAmount(profit)))
			}
		}
	}
	if len(netProfitLines) > 0 && hasIncomeTax {
		if stated, ok := sumReportLines(netProfitLines, layout, -1); ok {
			calculated := profit - incomeTax
			tolerance := math.Max(1, math.Abs(calculated)*0.01)
			if math.Abs(stated-calculated) > tolerance {
				addWarning(fmt.Sprintf("Строка «%s» показывает чистую прибыль %s, а расчёт прибыль до налога минус налог — %s.", netProfitLines[0].label, formatReportAmount(stated), formatReportAmount(calculated)))
			}
		}
	}
	if hasOutputVat && outputVat > grossRevenue {
		addWarning("Начисленный НДС больше выручки: проверьте строки НДС и выручки в книге.")
	}
	if netRevenue <= 0 {
		addWarning("Выручка без НДС равна нулю или отрицательна: маржинальность рассчитать нельзя.")
	}
	if !hasRevenue {
		addWarning("Не найдена строка выручки: итоги отчёта могут быть неполными.")
	}
	if !hasCost && !hasExpense {
		addWarning("Не найдены строки себестоимости и затрат: прибыль может быть завышена.")
	}
	if !hasOutputVat {
		addWarning("Строка начисленного НДС не найдена: выручка использована без корректировки на НДС.")
	}
	if !hasVatPayable {
		addWarning("Строка НДС к уплате не найдена: справочная сумма НДС не показана.")
	}
	if !hasIncomeTax && !hasNetProfit {
		addWarning("Строка налога на прибыль не найдена: чистая прибыль не рассчитана.")
	}
	if len(reportLinesWithRole(layout.lines, reportRoleVatInput)) > 0 {
		addWarning("В книге есть входящий НДС. Себестоимость и затраты использованы как записано в книге; если они включают входящий НДС, прибыль может быть занижена. Проверьте, что затраты указаны без НДС.")
	}
	if countReportRole(layout.lines, reportRoleRevenueTotal) > 1 {
		addWarning("Найдено несколько итоговых строк выручки: их значения сложены.")
	}
	if countReportRole(layout.lines, reportRoleCostTotal) > 1 {
		addWarning("Найдено несколько итоговых строк себестоимости: их значения сложены.")
	}
	if countReportRole(layout.lines, reportRoleExpenseTotal) > 1 {
		addWarning("Найдено несколько итоговых строк затрат: их значения сложены.")
	}
	if countReportRole(layout.lines, reportRoleIncomeTax) > 1 {
		addWarning("Найдено несколько строк налога на прибыль: их значения сложены.")
	}

	periods := buildReportPeriods(table, layout, revenueLines, outputVatLines, costLines, expenseLines, vatPayableLines, incomeTaxLines, netProfitLines)
	for _, period := range periods {
		if reportPeriodIsEmpty(period) {
			addWarning(fmt.Sprintf("Период «%s» полностью нулевой: возможно, данные за этот период не заполнены.", period.Label))
		}
	}

	if layout.totalColumn < 0 {
		addWarning("Итоговый столбец «Итого» не найден: суммы рассчитаны по периодам.")
	} else {
		checkReportReconciliation("выручка", revenueLines, layout, addWarning)
		checkReportReconciliation("начисленный НДС", outputVatLines, layout, addWarning)
		checkReportReconciliation("себестоимость", costLines, layout, addWarning)
		checkReportReconciliation("затраты", expenseLines, layout, addWarning)
		checkReportReconciliation("налог на прибыль", incomeTaxLines, layout, addWarning)
		checkReportReconciliation("НДС к уплате", vatPayableLines, layout, addWarning)
	}

	labels := make([]string, 0, len(periods))
	for _, period := range periods {
		labels = append(labels, period.Label)
	}
	totalHeader := reportHeader(table, layout, layout.totalColumn)
	if totalHeader == "" {
		totalHeader = "сумма периодов"
	}
	addWarning(fmt.Sprintf("Книга прочитана как отчёт о прибыли: статьи в столбце %s, периоды: %s; итог — «%s».", layout.excelColumn(layout.labelColumn), strings.Join(labels, ", "), totalHeader))
	if hasOutputVat {
		addWarning(fmt.Sprintf("Выручка без НДС рассчитана как «%s» минус начисленный НДС «%s».", firstReportLabel(revenueLines, "Выручка"), firstReportLabel(outputVatLines, "НДС")))
	}
	if hasCost || hasExpense {
		addWarning("Прибыль до налога на прибыль рассчитана как выручка без НДС минус себестоимость и затраты.")
	}
	if hasVatPayable {
		addWarning("НДС к уплате показан справочно и не вычитается из прибыли: это расчёт с бюджетом, а не расход периода.")
	}
	if hasIncomeTax {
		addWarning("Налог на прибыль взят из отчёта; чистая прибыль = прибыль до налога − налог на прибыль.")
	}
	addWarning("Значения формул взяты из сохранённых в книге результатов; приложение не пересчитывает формулы Excel.")
	addWarning("Это управленческая оценка по содержимому книги, а не бухгалтерская или налоговая отчётность.")

	reportLines := buildReportLines(layout, revenueLines, costLines, expenseLines, outputVatLines, incomeTaxLines, netProfitLines, vatPayableLines, grossRevenue, outputVat, netRevenue, totalCost, profit, incomeTax, netProfit, vatPayable, hasOutputVat, hasIncomeTax, hasNetProfit, hasVatPayable)

	columns := []ColumnMapping{{
		Role:        "Строки отчёта",
		Header:      reportHeaderOrFallback(table, layout, layout.labelColumn),
		ExcelColumn: layout.excelColumn(layout.labelColumn),
	}}
	for _, column := range layout.periodColumns() {
		columns = append(columns, ColumnMapping{
			Role:        "Период",
			Header:      reportHeader(table, layout, column),
			ExcelColumn: layout.excelColumn(column),
		})
	}
	if layout.totalColumn >= 0 {
		columns = append(columns, ColumnMapping{
			Role:        "Итоговый столбец",
			Header:      reportHeader(table, layout, layout.totalColumn),
			ExcelColumn: layout.excelColumn(layout.totalColumn),
		})
	}

	report := Report{
		Mode:              ModeReport,
		SheetName:         table.SheetName,
		TableName:         table.TableName,
		Range:             table.Range,
		StructuredTable:   table.Structured,
		HeaderRow:         layout.headerIndex + 1,
		HeaderExcelRow:    excelRowAt(table, layout.headerIndex),
		RowsInTable:       len(table.Rows) - layout.headerIndex - 1,
		RowsAnalyzed:      len(layout.lines),
		RowsWithRevenue:   boolInt(hasRevenue),
		RowsWithCost:      boolInt(hasCost || hasExpense),
		RowsWithProfit:    1,
		Columns:           columns,
		ReportLines:       reportLines,
		Periods:           periods,
		TopItems:          []LineItem{},
		LowestItems:       []LineItem{},
		Warnings:          []string{},
		TotalColumnHeader: totalHeader,
	}
	if hasRevenue {
		report.RevenueTotal = floatPointer(netRevenue)
	}
	if hasCost || hasExpense {
		report.CostTotal = floatPointer(totalCost)
	}
	report.ProfitTotal = floatPointer(profit)
	if netRevenue != 0 {
		report.MarginPercent = floatPointer(profit / netRevenue * 100)
	}
	report.GrossRevenue = floatPointer(grossRevenue)
	if hasOutputVat {
		report.OutputVat = floatPointer(outputVat)
	}
	if hasVatPayable {
		report.VatPayable = floatPointer(vatPayable)
	}
	if hasIncomeTax {
		report.IncomeTax = floatPointer(incomeTax)
	}
	if hasNetProfit {
		report.NetProfit = floatPointer(netProfit)
		if netRevenue != 0 {
			report.NetMarginPercent = floatPointer(netProfit / netRevenue * 100)
		}
	}
	if len(periods) > 0 {
		for _, period := range periods {
			if period.Profit == nil {
				continue
			}
			if *period.Profit < 0 {
				report.LossRows++
			} else if *period.Profit > 0 {
				report.ProfitableRows++
			}
		}
	} else if profit < 0 {
		report.LossRows = 1
	} else if profit > 0 {
		report.ProfitableRows = 1
	}

	for warning := range warnings {
		report.Warnings = append(report.Warnings, warning)
	}
	sort.Strings(report.Warnings)
	return report
}

func buildReportPeriods(table xlsx.Table, layout reportLayout, revenueLines, outputVatLines, costLines, expenseLines, vatPayableLines, incomeTaxLines, netProfitLines []reportSourceLine) []ReportPeriod {
	periods := make([]ReportPeriod, 0, len(layout.periodColumns()))
	for _, column := range layout.periodColumns() {
		grossRevenue, hasRevenue := sumReportLines(revenueLines, layout, column)
		outputVat, hasOutputVat := sumReportLines(outputVatLines, layout, column)
		revenue := grossRevenue
		if hasOutputVat {
			revenue -= outputVat
		}
		cost, _ := sumReportLines(costLines, layout, column)
		expense, _ := sumReportLines(expenseLines, layout, column)
		totalCost := cost + expense
		profit := revenue - totalCost
		incomeTax, hasIncomeTax := sumReportLines(incomeTaxLines, layout, column)
		netProfit := profit
		hasNetProfit := false
		if len(netProfitLines) > 0 {
			if value, ok := sumReportLines(netProfitLines, layout, column); ok {
				netProfit = value
				hasNetProfit = true
			}
		}
		if !hasNetProfit && hasIncomeTax {
			netProfit = profit - incomeTax
			hasNetProfit = true
		}
		vatPayable, hasVatPayable := sumReportLines(vatPayableLines, layout, column)

		period := ReportPeriod{
			Label:       reportHeader(table, layout, column),
			ExcelColumn: layout.excelColumn(column),
		}
		if hasRevenue {
			period.GrossRevenue = floatPointer(grossRevenue)
			period.Revenue = floatPointer(revenue)
		}
		if len(costLines) > 0 || len(expenseLines) > 0 {
			period.Cost = floatPointer(totalCost)
		}
		period.Profit = floatPointer(profit)
		if revenue != 0 {
			period.MarginPercent = floatPointer(profit / revenue * 100)
		}
		if hasIncomeTax {
			period.IncomeTax = floatPointer(incomeTax)
		}
		if hasNetProfit {
			period.NetProfit = floatPointer(netProfit)
			if revenue != 0 {
				period.NetMarginPercent = floatPointer(netProfit / revenue * 100)
			}
		}
		if hasVatPayable {
			period.VatPayable = floatPointer(vatPayable)
		}
		periods = append(periods, period)
	}
	return periods
}

func buildReportLines(layout reportLayout, revenueLines, costLines, expenseLines, outputVatLines, incomeTaxLines, netProfitLines, vatPayableLines []reportSourceLine, grossRevenue, outputVat, netRevenue, totalCost, profit, incomeTax, netProfit, vatPayable float64, hasOutputVat, hasIncomeTax, hasNetProfit, hasVatPayable bool) []ReportLine {
	lines := make([]ReportLine, 0, 12)
	appendLine := func(label, role string, amount float64, excelRow int, subtotal bool) {
		lines = append(lines, ReportLine{
			Label:    label,
			Role:     role,
			ExcelRow: excelRow,
			Amount:   floatPointer(amount),
			Subtotal: subtotal,
		})
	}

	if hasOutputVat {
		appendLine(firstReportLabel(revenueLines, "Итого выручка"), "Выручка с НДС", grossRevenue, firstReportExcelRow(revenueLines), false)
		appendLine(firstReportLabel(outputVatLines, "Начисленный НДС"), "НДС начисленный", -outputVat, firstReportExcelRow(outputVatLines), false)
		appendLine("Выручка без НДС", "Выручка", netRevenue, 0, true)
	} else {
		appendLine(firstReportLabel(revenueLines, "Выручка"), "Выручка", grossRevenue, firstReportExcelRow(revenueLines), false)
	}

	for _, line := range costLines {
		amount, _ := sumReportLines([]reportSourceLine{line}, layout, -1)
		appendLine(line.label, "Себестоимость", -amount, line.excelRow, false)
	}
	for _, line := range expenseLines {
		amount, _ := sumReportLines([]reportSourceLine{line}, layout, -1)
		appendLine(line.label, "Затраты", -amount, line.excelRow, false)
	}
	if len(costLines)+len(expenseLines) > 1 {
		appendLine("Себестоимость и затраты, всего", "Расходы", -totalCost, 0, true)
	}

	appendLine("Прибыль до налога на прибыль", "Прибыль", profit, 0, true)
	if hasIncomeTax {
		for _, line := range incomeTaxLines {
			amount, _ := sumReportLines([]reportSourceLine{line}, layout, -1)
			appendLine(line.label, "Налог на прибыль", -amount, line.excelRow, false)
		}
	}
	if hasNetProfit {
		appendLine(firstReportLabel(netProfitLines, "Чистая прибыль"), "Чистая прибыль", netProfit, firstReportExcelRow(netProfitLines), true)
	}
	if hasVatPayable {
		appendLine("НДС к уплате (справочно)", "Справочно", vatPayable, firstReportExcelRow(vatPayableLines), false)
	}
	return lines
}

func reportPeriodIsEmpty(period ReportPeriod) bool {
	values := []*float64{
		period.GrossRevenue,
		period.Revenue,
		period.Cost,
		period.Profit,
		period.IncomeTax,
		period.NetProfit,
		period.VatPayable,
	}
	for _, value := range values {
		if value != nil && math.Abs(*value) > 0.000001 {
			return false
		}
	}
	return true
}

func checkReportReconciliation(name string, lines []reportSourceLine, layout reportLayout, addWarning func(string)) {
	if len(lines) == 0 || layout.totalColumn < 0 {
		return
	}
	total, _ := sumReportLines(lines, layout, -1)
	periodSum := 0.0
	for _, period := range layout.periodColumns() {
		value, _ := sumReportLines(lines, layout, period)
		periodSum += value
	}
	if math.Abs(total-periodSum) > 1 {
		addWarning(fmt.Sprintf("Итог по статье «%s» (%s) расходится с суммой периодов (%s). Проверьте формулы книги.", name, formatReportAmount(total), formatReportAmount(periodSum)))
	}
}

func sumReportLines(lines []reportSourceLine, layout reportLayout, column int) (float64, bool) {
	if len(lines) == 0 {
		return 0, false
	}
	total := 0.0
	if column >= 0 {
		for _, line := range lines {
			total += line.values[column]
		}
		return total, true
	}
	for _, line := range lines {
		if layout.totalColumn >= 0 {
			if value, ok := line.values[layout.totalColumn]; ok {
				total += value
				continue
			}
		}
		for _, period := range layout.periodColumns() {
			total += line.values[period]
		}
	}
	return total, true
}

func selectReportLines(lines []reportSourceLine, totalRole, componentRole reportRole) []reportSourceLine {
	if totals := reportLinesWithRole(lines, totalRole); len(totals) > 0 {
		return totals
	}
	return reportLinesWithRole(lines, componentRole)
}

func reportLinesWithRole(lines []reportSourceLine, role reportRole) []reportSourceLine {
	var result []reportSourceLine
	for _, line := range lines {
		if line.role == role {
			result = append(result, line)
		}
	}
	return result
}

func countReportRole(lines []reportSourceLine, role reportRole) int {
	count := 0
	for _, line := range lines {
		if line.role == role {
			count++
		}
	}
	return count
}

func reportCell(row []string, column int) string {
	if column < 0 || column >= len(row) {
		return ""
	}
	return row[column]
}

func reportHeader(table xlsx.Table, layout reportLayout, column int) string {
	if column < 0 || layout.headerIndex < 0 || layout.headerIndex >= len(table.Rows) {
		return ""
	}
	return strings.TrimSpace(reportCell(table.Rows[layout.headerIndex], column))
}

func reportHeaderOrFallback(table xlsx.Table, layout reportLayout, column int) string {
	if header := reportHeader(table, layout, column); header != "" {
		return header
	}
	return "Столбец " + layout.excelColumn(column)
}

func firstReportLabel(lines []reportSourceLine, fallback string) string {
	if len(lines) == 0 {
		return fallback
	}
	return lines[0].label
}

func firstReportExcelRow(lines []reportSourceLine) int {
	if len(lines) == 0 {
		return 0
	}
	return lines[0].excelRow
}

func containsReportColumn(columns []int, target int) bool {
	for _, column := range columns {
		if column == target {
			return true
		}
	}
	return false
}

func boolInt(value bool) int {
	if value {
		return 1
	}
	return 0
}

func formatReportAmount(value float64) string {
	return strings.ReplaceAll(fmt.Sprintf("%.2f", value), ".", ",")
}

func tableStartColumn(reference string) int {
	reference = strings.TrimSpace(reference)
	if bang := strings.LastIndex(reference, "!"); bang >= 0 {
		reference = reference[bang+1:]
	}
	if colon := strings.Index(reference, ":"); colon >= 0 {
		reference = reference[:colon]
	}
	column := 0
	for _, character := range reference {
		if character >= 'a' && character <= 'z' {
			character -= 'a' - 'A'
		}
		if character < 'A' || character > 'Z' {
			break
		}
		column = column*26 + int(character-'A'+1)
	}
	if column == 0 {
		return 1
	}
	return column
}
