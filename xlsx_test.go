// Copyright (c) 2018 Beta Kuang
//
// This software is released under the MIT License.
// https://opensource.org/licenses/MIT

package spreadsheet_test

import (
	"testing"

	"github.com/beta/spreadsheet"
)

func TestReadXLSXSheets(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	sheetCount := len(ss.Sheets)
	if sheetCount != 2 {
		t.Errorf("incorrect sheet count, want %d, got %d", 2, sheetCount)
	}
}

func TestReadXLSXSheetByName(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.SheetsByName["Sheet 1"]
	if s0 != ss.Sheets[0] {
		t.Errorf("\"Sheet 1\" should be equivalent to the first sheet")
	}

	s1 := ss.SheetsByName["Sheet 2"]
	if s1 != ss.Sheets[1] {
		t.Errorf("\"Sheet 2\" should be equivalent to the second sheet")
	}
}

func TestReadXLSXRowCells(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.Sheets[0]

	rowCount := len(s0.Rows)
	if rowCount != 6 {
		t.Errorf("incorrect row count, want %d, got %d", 6, rowCount)
	}

	cellCount := len(s0.Rows[0].Cells)
	if cellCount != 3 {
		t.Errorf("incorrect cell count, want %d, got %d", 3, cellCount)
	}
}

func TestReadXLSXNumCell(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.Sheets[0]

	cell := s0.Rows[1].Cells[0]
	cellType := cell.Type()
	if cellType != spreadsheet.Numeric {
		t.Errorf("incorrect cell data type, want %s, got %s",
			spreadsheet.Numeric.Name(), cellType.Name())
	}

	intValue, err := cell.Int()
	if err != nil {
		t.Errorf("failed to get cell value, error: %v", err)
	}
	if intValue != 1 {
		t.Errorf("incorrect cell value, want %d, got %d", 1, intValue)
	}

	i64Value, err := cell.Int64()
	if err != nil {
		t.Errorf("failed to get cell value, error: %v", err)
	}
	if i64Value != 1 {
		t.Errorf("incorrect cell value, want %d, got %d", 1, i64Value)
	}

	floatValue, err := cell.Float()
	if err != nil {
		t.Errorf("failed to get cell value, error: %v", err)
	}
	if !(floatValue > 0.9 && floatValue < 1.1) {
		t.Errorf("incorrect cell value, want %f, got %f", 1., floatValue)
	}
}

func TestReadXLSXBoolCell(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.Sheets[0]

	cell := s0.Rows[4].Cells[2]
	cellType := cell.Type()
	if !cell.Is(spreadsheet.Bool) {
		t.Errorf("incorrect cell data type, want %s, got %s",
			spreadsheet.Bool.Name(), cellType.Name())
	}

	boolValue, err := cell.Bool()
	if err != nil {
		t.Errorf("failed to get cell value, error: %v", err)
	}
	if boolValue != true {
		t.Errorf("incorrect cell value, want %t, got %t", true, boolValue)
	}
}

func TestReadXLSXStringCell(t *testing.T) {
	ss, err := spreadsheet.Open("test.xlsx")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.Sheets[0]

	cell := s0.Rows[5].Cells[1]
	cellType := cell.Type()
	if !cell.Is(spreadsheet.String) {
		t.Errorf("incorrect cell data type, want %s, got %s",
			spreadsheet.String.Name(), cellType.Name())
	}

	strValue := cell.String()
	if strValue != "e" {
		t.Errorf("incorrect cell value, want %s, got %s", "e", strValue)
	}
}
