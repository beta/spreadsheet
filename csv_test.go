// Copyright (c) 2018 Beta Kuang
//
// This software is released under the MIT License.
// https://opensource.org/licenses/MIT

package spreadsheet_test

import (
	"strings"
	"testing"

	"github.com/beta/spreadsheet"
)

func TestReadCSVSheets(t *testing.T) {
	ss, err := spreadsheet.Open("test.csv")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	sheetCount := len(ss.Sheets)
	if sheetCount != 1 {
		t.Errorf("incorrect sheet count, want %d, got %d", 1, sheetCount)
	}
}

func TestReadCSVSheetByName(t *testing.T) {
	ss, err := spreadsheet.Open("test.csv")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s0 := ss.SheetsByName["Sheet 1"]
	if s0 != ss.Sheets[0] {
		t.Errorf("\"Sheet 1\" should be equivalent to the first sheet")
	}
}

func TestReadCSVRowCells(t *testing.T) {
	ss, err := spreadsheet.Open("test.csv")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s := ss.Sheets[0]

	rowCount := len(s.Rows)
	if rowCount != 6 {
		t.Errorf("incorrect row count, want %d, got %d", 6, rowCount)
	}

	cellCount := len(s.Rows[0].Cells)
	if cellCount != 3 {
		t.Errorf("incorrect cell count, want %d, got %d", 3, cellCount)
	}
}

func TestReadCSVCellValue(t *testing.T) {
	ss, err := spreadsheet.Open("test.csv")
	if err != nil {
		t.Errorf("failed to open file, error: %v", err)
	}

	s := ss.Sheets[0]

	cell := s.Rows[0].Cells[1]
	cellType := cell.Type()
	if cellType != spreadsheet.String {
		t.Errorf("incorrect cell type, want %s, got %s",
			spreadsheet.String.Name(), cellType.Name())
	}
	cellValue := cell.String()
	if cellValue != "name" {
		t.Errorf("incorrect cell value, want %s, got %s", "name", cellValue)
	}

	commaCell := s.Cell(3, 2)
	if !commaCell.Is(spreadsheet.String) {
		t.Errorf("incorrect cell type, want %s, got %s",
			spreadsheet.String.Name(), commaCell.Type().Name())
	}
	commaCellValue := commaCell.String()
	if !strings.Contains(commaCellValue, ",") {
		t.Errorf("missing comma in cell value, got %s", commaCellValue)
	}
}
