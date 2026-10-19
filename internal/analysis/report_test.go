package analysis

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/slchumak75-sketch/analitika-sales/internal/xlsx"
)

func TestAnalyzeReportWithVatAndIncomeTax(t *testing.T) {
	table := xlsx.Table{
		SheetName: "Лист1",
		Range:     "A1:F8",
		Rows: [][]string{
			{"Статья", "1кв", "2кв", "3кв", "4кв", "Итого"},
			{"Налог НДС", "50", "100", "150", "0", "300"},
			{"Налог на прибыль", "40", "80", "120", "0", "240"},
			{"Итого выручка", "1000", "2000", "3000", "0", "6000"},
			{"Итого НДС", "200", "400", "600", "0", "1200"},
			{"Себестоимость реализация", "500", "1000", "1500", "0", "3000"},
			{"Себестоимость затрат", "100", "200", "300", "0", "600"},
			{"Входящий НДС", "100", "200", "300", "0", "600"},
		},
		ExcelRows: []int{1, 2, 3, 4, 5, 6, 7, 8},
	}

	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	if report.Mode != ModeReport {
		t.Fatalf("mode = %q, want %q", report.Mode, ModeReport)
	}
	if report.HeaderExcelRow != 1 || report.RowsInTable != 7 {
		t.Fatalf("unexpected header/row counts: header=%d rows=%d", report.HeaderExcelRow, report.RowsInTable)
	}
	assertNear(t, value(report.GrossRevenue), 6000)
	assertNear(t, value(report.OutputVat), 1200)
	assertNear(t, value(report.RevenueTotal), 4800)
	assertNear(t, value(report.CostTotal), 3600)
	assertNear(t, value(report.ProfitTotal), 1200)
	assertNear(t, value(report.MarginPercent), 25)
	assertNear(t, value(report.IncomeTax), 240)
	assertNear(t, value(report.NetProfit), 960)
	assertNear(t, value(report.NetMarginPercent), 20)
	assertNear(t, value(report.VatPayable), 300)
	if report.TotalColumnHeader != "Итого" {
		t.Fatalf("total column header = %q", report.TotalColumnHeader)
	}
	if len(report.Periods) != 4 {
		t.Fatalf("periods = %d, want 4", len(report.Periods))
	}
	first := report.Periods[0]
	if first.Label != "1кв" || first.ExcelColumn != "B" {
		t.Fatalf("unexpected first period: %+v", first)
	}
	assertNear(t, value(first.GrossRevenue), 1000)
	assertNear(t, value(first.Revenue), 800)
	assertNear(t, value(first.Cost), 600)
	assertNear(t, value(first.Profit), 200)
	assertNear(t, value(first.MarginPercent), 25)
	assertNear(t, value(first.IncomeTax), 40)
	assertNear(t, value(first.NetProfit), 160)
	assertNear(t, value(first.NetMarginPercent), 20)
	assertNear(t, value(first.VatPayable), 50)
	last := report.Periods[3]
	if last.Label != "4кв" || last.MarginPercent != nil || last.NetMarginPercent != nil {
		t.Fatalf("unexpected zero period: %+v", last)
	}
	assertNear(t, value(last.Profit), 0)
	if !hasReportWarning(report.Warnings, "4кв") {
		t.Fatalf("expected a warning about the empty fourth quarter, got %v", report.Warnings)
	}
	if !reportHasLine(report.ReportLines, "Выручка без НДС") || !reportHasLine(report.ReportLines, "Чистая прибыль") {
		t.Fatalf("unexpected report lines: %+v", report.ReportLines)
	}
	if len(report.Columns) != 6 || report.Columns[0].ExcelColumn != "A" || report.Columns[len(report.Columns)-1].ExcelColumn != "F" {
		t.Fatalf("unexpected column mappings: %+v", report.Columns)
	}
}

func TestAnalyzeReportWithoutVat(t *testing.T) {
	table := xlsx.Table{
		SheetName: "P&L",
		Range:     "A1:C4",
		Rows: [][]string{
			{"Статья", "1кв", "Итого"},
			{"Выручка", "1000", "1000"},
			{"Себестоимость", "600", "600"},
			{"Налог на прибыль", "80", "80"},
		},
		ExcelRows: []int{1, 2, 3, 4},
	}
	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	if report.Mode != ModeReport {
		t.Fatalf("mode = %q, want %q", report.Mode, ModeReport)
	}
	if report.OutputVat != nil {
		t.Fatalf("output VAT should be absent, got %v", value(report.OutputVat))
	}
	assertNear(t, value(report.GrossRevenue), 1000)
	assertNear(t, value(report.RevenueTotal), 1000)
	assertNear(t, value(report.CostTotal), 600)
	assertNear(t, value(report.ProfitTotal), 400)
	assertNear(t, value(report.MarginPercent), 40)
	assertNear(t, value(report.IncomeTax), 80)
	assertNear(t, value(report.NetProfit), 320)
	assertNear(t, value(report.NetMarginPercent), 32)
	if report.VatPayable != nil {
		t.Fatalf("VAT payable should be absent, got %v", value(report.VatPayable))
	}
	if len(report.Periods) != 1 || report.Periods[0].Label != "1кв" {
		t.Fatalf("unexpected periods: %+v", report.Periods)
	}
	if !hasReportWarning(report.Warnings, "начисленного НДС") {
		t.Fatalf("expected a missing-VAT warning, got %v", report.Warnings)
	}
	if reportHasLine(report.ReportLines, "НДС к уплате (справочно)") {
		t.Fatalf("unexpected VAT payable line: %+v", report.ReportLines)
	}
}

func TestClassifyReportLabel(t *testing.T) {
	tests := []struct {
		label string
		want  reportRole
	}{
		{"Итого выручка", reportRoleRevenueTotal},
		{"Выручка по отчетам ККМ", reportRoleRevenueComponent},
		{"Реализация (Накладные)", reportRoleRevenueComponent},
		{"Налог НДС", reportRoleVatPayable},
		{"Итого НДС", reportRoleOutputVatTotal},
		{"НДС Выручка по отчетам ККМ", reportRoleOutputVatComponent},
		{"НДС к уплате", reportRoleVatPayable},
		{"Входящий НДС с вычетом", reportRoleVatInput},
		{"Вычет по товарам без НДС", reportRoleNone},
		{"Товары реализация без НДС", reportRoleNone},
		{"Себестоимость реализация", reportRoleCost},
		{"Себестоимость затрат", reportRoleExpense},
		{"Итого расходы", reportRoleExpenseTotal},
		{"Налог на прибыль", reportRoleIncomeTax},
		{"Чистая прибыль", reportRoleNetProfit},
		{"Прибыль до налогообложения", reportRoleProfitBeforeTax},
		{"фонд з/пл", reportRoleNone},
	}
	for _, test := range tests {
		if got := classifyReportLabel(test.label); got != test.want {
			t.Errorf("classifyReportLabel(%q) = %v, want %v", test.label, got, test.want)
		}
	}
}

func TestAnalyzeItemTableWithTotalColumnUsesItemMode(t *testing.T) {
	table := xlsx.Table{
		SheetName: "Продажи",
		Range:     "A1:D3",
		HasHeader: true,
		Rows: [][]string{
			{"Товар", "Выручка", "Себестоимость", "Итого"},
			{"Кофе", "100", "60", "40"},
			{"Чай", "200", "150", "50"},
		},
		ExcelRows: []int{1, 2, 3},
	}
	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	if report.Mode != ModeItems {
		t.Fatalf("mode = %q, want %q", report.Mode, ModeItems)
	}
	assertNear(t, value(report.RevenueTotal), 300)
	assertNear(t, value(report.CostTotal), 210)
	assertNear(t, value(report.ProfitTotal), 90)
	assertNear(t, value(report.MarginPercent), 30)
}

func TestAnalyzeRepositoryWorkbook(t *testing.T) {
	path := filepath.Join("..", "..", "ОТЧЕТ ПО ОРГАНИЗАЦИИ мир.xlsx")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Skipf("repository workbook is not available: %v", err)
	}
	table, err := xlsx.ReadFirstTable(data)
	if err != nil {
		t.Fatalf("ReadFirstTable() error = %v", err)
	}
	if table.SheetName != "Лист1" || table.Range != "B2:H45" {
		t.Fatalf("unexpected table: sheet=%q range=%q", table.SheetName, table.Range)
	}
	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	if report.Mode != ModeReport {
		t.Fatalf("mode = %q, want %q", report.Mode, ModeReport)
	}
	assertNear(t, value(report.GrossRevenue), 1468005.41)
	assertNear(t, value(report.OutputVat), 239877.9033333333)
	assertNear(t, value(report.RevenueTotal), 1228127.5066666666)
	assertNear(t, value(report.CostTotal), 1111424.21)
	assertNear(t, value(report.ProfitTotal), 116703.29666666663)
	assertNear(t, value(report.MarginPercent), 9.50253911203552)
	assertNear(t, value(report.IncomeTax), 23340.659333333322)
	assertNear(t, value(report.NetProfit), 93362.63733333332)
	assertNear(t, value(report.NetMarginPercent), 7.602031289628417)
	assertNear(t, value(report.VatPayable), 57001.7533333333)
	if report.TotalColumnHeader != "Итого" || len(report.Periods) != 4 {
		t.Fatalf("unexpected report metadata: total=%q periods=%d", report.TotalColumnHeader, len(report.Periods))
	}
	if report.Columns[0].ExcelColumn != "B" || report.Columns[len(report.Columns)-1].ExcelColumn != "H" {
		t.Fatalf("unexpected column mappings: %+v", report.Columns)
	}
	first := report.Periods[0]
	if first.Label != "1кв" {
		t.Fatalf("first period = %q", first.Label)
	}
	assertNear(t, value(first.GrossRevenue), 258823.44)
	assertNear(t, value(first.Revenue), 216963.11666666667)
	assertNear(t, value(first.Cost), 214232.24000000002)
	assertNear(t, value(first.Profit), 2730.876666666649)
	assertNear(t, value(first.MarginPercent), 1.2586824473314775)
	assertNear(t, value(first.IncomeTax), 546.1753333333312)
	assertNear(t, value(first.NetProfit), 2184.7013333333175)
	assertNear(t, value(first.NetMarginPercent), 1.0069459578651814)
	last := report.Periods[3]
	if last.Label != "4кв" || last.MarginPercent != nil || last.NetMarginPercent != nil {
		t.Fatalf("unexpected fourth quarter: %+v", last)
	}
	assertNear(t, value(last.GrossRevenue), 0)
	assertNear(t, value(last.Profit), 0)
	if !hasReportWarning(report.Warnings, "4кв") || !hasReportWarning(report.Warnings, "входящий НДС") {
		t.Fatalf("expected quarter and VAT warnings, got %v", report.Warnings)
	}
}

func reportHasLine(lines []ReportLine, label string) bool {
	for _, line := range lines {
		if line.Label == label {
			return true
		}
	}
	return false
}

func hasReportWarning(warnings []string, substring string) bool {
	for _, warning := range warnings {
		if strings.Contains(warning, substring) {
			return true
		}
	}
	return false
}
