package xlsx

import (
	"archive/zip"
	"bytes"
	"testing"
)

func TestReadFirstStructuredTable(t *testing.T) {
	workbook := testWorkbook(t, map[string]string{
		"xl/workbook.xml":                     `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Продажи" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels":          `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		"xl/sharedStrings.xml":                `<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Товар</t></si><si><t>Выручка</t></si><si><t>Позиция А</t></si><si><t>Позиция Б</t></si></sst>`,
		"xl/worksheets/sheet1.xml":            `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Выручка</t></is></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2"><v>1250.5</v></c></row><row r="3"><c r="A3" t="s"><v>3</v></c><c r="B3"><v>800</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId2"/></tableParts></worksheet>`,
		"xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`,
		"xl/tables/table1.xml":                `<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Sales" displayName="SalesTable" ref="A1:B3" headerRowCount="1"><tableColumns count="2"><tableColumn id="1" name="Товар"/><tableColumn id="2" name="Выручка"/></tableColumns></table>`,
	})

	table, err := ReadFirstTable(workbook)
	if err != nil {
		t.Fatalf("ReadFirstTable() error = %v", err)
	}
	if !table.Structured {
		t.Fatal("expected a structured Excel table")
	}
	if table.SheetName != "Продажи" || table.TableName != "SalesTable" || table.Range != "A1:B3" {
		t.Fatalf("unexpected table metadata: %+v", table)
	}
	if !table.HasHeader || len(table.Rows) != 3 || table.Rows[1][0] != "Позиция А" || table.Rows[1][1] != "1250.5" {
		t.Fatalf("unexpected table content: %+v", table.Rows)
	}
	if table.ExcelRows[0] != 1 || table.ExcelRows[2] != 3 {
		t.Fatalf("unexpected Excel row references: %v", table.ExcelRows)
	}
}

func TestReadFirstNonEmptySheetWhenNoStructuredTable(t *testing.T) {
	workbook := testWorkbook(t, map[string]string{
		"xl/workbook.xml":            `<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Отчёт" sheetId="1" r:id="rId1"/><sheet name="Пустой" sheetId="2" r:id="rId2"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<?xml version="1.0"?><Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>`,
		"xl/worksheets/sheet1.xml":   `<?xml version="1.0"?><worksheet><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>Позиция</t></is></c><c r="C2"><v>12</v></c></row><row r="3"><c r="B3" t="inlineStr"><is><t>Ещё</t></is></c><c r="C3"><v>8</v></c></row></sheetData></worksheet>`,
		"xl/worksheets/sheet2.xml":   `<?xml version="1.0"?><worksheet><sheetData/></worksheet>`,
	})

	table, err := ReadFirstTable(workbook)
	if err != nil {
		t.Fatalf("ReadFirstTable() error = %v", err)
	}
	if table.Structured || table.SheetName != "Отчёт" || table.Range != "B2:C3" {
		t.Fatalf("unexpected fallback table: %+v", table)
	}
	if table.Rows[0][0] != "Позиция" || table.Rows[0][1] != "12" {
		t.Fatalf("unexpected fallback rows: %+v", table.Rows)
	}
}

func TestReadFirstTableRejectsInvalidZip(t *testing.T) {
	if _, err := ReadFirstTable([]byte("not an xlsx file")); err == nil {
		t.Fatal("expected invalid XLSX input to return an error")
	}
}

func TestParseReferences(t *testing.T) {
	row, column, ok := parseCellReference("$AB$19")
	if !ok || row != 19 || column != 28 {
		t.Fatalf("parseCellReference = (%d, %d, %v), want (19, 28, true)", row, column, ok)
	}
	minColumn, minRow, maxColumn, maxRow, ok := parseRange("$B$2:$D$9")
	if !ok || minColumn != 2 || minRow != 2 || maxColumn != 4 || maxRow != 9 {
		t.Fatalf("parseRange = (%d, %d, %d, %d, %v)", minColumn, minRow, maxColumn, maxRow, ok)
	}
}

func testWorkbook(t *testing.T, files map[string]string) []byte {
	t.Helper()
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for name, contents := range files {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatalf("create zip entry %s: %v", name, err)
		}
		if _, err := entry.Write([]byte(contents)); err != nil {
			t.Fatalf("write zip entry %s: %v", name, err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("close test workbook: %v", err)
	}
	return output.Bytes()
}
