package analysis

import (
	"math"
	"testing"

	"github.com/slchumak75-sketch/analitika-sales/internal/xlsx"
)

func TestAnalyzeCalculatesProfitAndMargin(t *testing.T) {
	table := xlsx.Table{
		SheetName:  "Продажи",
		TableName:  "SalesTable",
		Range:      "A1:D4",
		HasHeader:  true,
		Structured: true,
		Rows: [][]string{
			{"Товар", "Выручка, BYN", "Себестоимость, BYN", "Количество"},
			{"Позиция A", "1 000,00", "700,00", "10"},
			{"Позиция B", "500,00", "650,00", "5"},
			{"Итого", "1500,00", "1350,00", "15"},
		},
		ExcelRows: []int{1, 2, 3, 4},
	}

	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	assertNear(t, value(report.RevenueTotal), 1500)
	assertNear(t, value(report.CostTotal), 1350)
	assertNear(t, value(report.ProfitTotal), 150)
	assertNear(t, value(report.MarginPercent), 10)
	if report.RowsAnalyzed != 2 || report.LossRows != 1 || report.ProfitableRows != 1 || report.TotalRowsSkipped != 1 {
		t.Fatalf("unexpected row counts: %+v", report)
	}
	if len(report.TopItems) == 0 || report.TopItems[0].Label != "Позиция A" || report.TopItems[0].Profit != 300 {
		t.Fatalf("unexpected top item: %+v", report.TopItems)
	}
	if len(report.LowestItems) == 0 || report.LowestItems[0].Label != "Позиция B" || report.LowestItems[0].Profit != -150 {
		t.Fatalf("unexpected lowest item: %+v", report.LowestItems)
	}
}

func TestAnalyzeUsesDirectProfitAndInfersCost(t *testing.T) {
	table := xlsx.Table{
		SheetName: "Лист1",
		Range:     "A1:C3",
		HasHeader: true,
		Rows: [][]string{
			{"Наименование товара", "Сумма продаж", "Прибыль"},
			{"Кофе", "200", "55"},
			{"Чай", "100", "-10"},
		},
		ExcelRows: []int{1, 2, 3},
	}

	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	assertNear(t, value(report.RevenueTotal), 300)
	assertNear(t, value(report.ProfitTotal), 45)
	assertNear(t, value(report.CostTotal), 255)
	assertNear(t, value(report.MarginPercent), 15)
	if report.RowsWithCost != 2 {
		t.Fatalf("expected costs to be inferred for both rows, got %d", report.RowsWithCost)
	}
}

func TestAnalyzeFindsHeaderAfterTitle(t *testing.T) {
	table := xlsx.Table{
		SheetName: "Отчёт",
		Range:     "A1:C4",
		Rows: [][]string{
			{"Отчёт о продажах", "", ""},
			{"Товар", "Выручка", "Себестоимость"},
			{"Папка", "80", "50"},
			{"Книга", "90", "40"},
		},
		ExcelRows: []int{1, 2, 3, 4},
	}

	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	if report.HeaderExcelRow != 2 || report.RowsAnalyzed != 2 {
		t.Fatalf("expected header on Excel row 2 and two data rows, got row=%d rows=%d", report.HeaderExcelRow, report.RowsAnalyzed)
	}
	assertNear(t, value(report.ProfitTotal), 80)
}

func TestAnalyzeUsesMarginPercentColumn(t *testing.T) {
	table := xlsx.Table{
		SheetName: "Продажи",
		Range:     "A1:C2",
		HasHeader: true,
		Rows: [][]string{
			{"Товар", "Выручка", "Маржа, %"},
			{"Позиция", "200", "25%"},
		},
		ExcelRows: []int{1, 2},
	}
	report, err := Analyze(table)
	if err != nil {
		t.Fatalf("Analyze() error = %v", err)
	}
	assertNear(t, value(report.ProfitTotal), 50)
	assertNear(t, value(report.MarginPercent), 25)
}

func TestParseNumberSupportsRussianMoneyAndNegatives(t *testing.T) {
	tests := []struct {
		input string
		want  float64
	}{
		{"1 234,56 руб.", 1234.56},
		{"(2 500,00)", -2500},
		{"$1,234.50", 1234.5},
		{"1,234,567", 1234567},
		{"-75.25", -75.25},
	}
	for _, test := range tests {
		got, ok := parseNumber(test.input)
		if !ok || math.Abs(got-test.want) > 0.000001 {
			t.Errorf("parseNumber(%q) = %v, %v; want %v, true", test.input, got, ok, test.want)
		}
	}
	if _, ok := parseNumber("—"); ok {
		t.Error("parseNumber(—) should be missing")
	}
}

func TestAnalyzeRejectsUnrecognizedFinancialHeaders(t *testing.T) {
	_, err := Analyze(xlsx.Table{
		HasHeader: true,
		Rows:      [][]string{{"Товар", "Комментарий"}, {"Позиция", "Обычный текст"}},
	})
	if err == nil {
		t.Fatal("expected an error for a table with no recognized financial columns")
	}
}

func value(pointer *float64) float64 {
	if pointer == nil {
		return math.NaN()
	}
	return *pointer
}

func assertNear(t *testing.T, got, want float64) {
	t.Helper()
	if math.IsNaN(got) || math.Abs(got-want) > 0.000001 {
		t.Fatalf("got %v, want %v", got, want)
	}
}
