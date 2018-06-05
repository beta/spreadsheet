// Copyright (c) 2018 Beta Kuang
//
// This software is released under the MIT License.
// https://opensource.org/licenses/MIT

// Package spreadsheet provides simple a interface to read spreadsheet files
// including XLSX and CSV.
package spreadsheet

import (
	"fmt"
	"path/filepath"
	"strconv"
	"strings"

	"github.com/pkg/errors"
)

// Open opens a spreadsheet file.
// filePath is the path to the spreadsheet file to be opened. Supported file
// types are XLSX and CSV (as described in RFC 4180).
// A Spreadsheet pointer will be returned if successful. Otherwise, a nil
// pointer will be returned with an error indicating what is wrong.
func Open(filePath string) (*Spreadsheet, error) {
	ext := filepath.Ext(filePath)
	switch strings.ToLower(ext[1:]) {
	case "csv":
		return openCSVFile(filePath)
	case "xlsx":
		return openXLSXFile(filePath)
	default:
		return nil, fmt.Errorf("input file is not a valid spreadsheet file")
	}
}

// Spreadsheet is a high-level interface for spreadsheet files including XLSX,
// XLS and CSV. A Spreadsheet contains one or more sheets of data.
type Spreadsheet struct {
	Sheets       []*Sheet // For CSV there's only 1 sheet, named "Sheet 1".
	SheetsByName map[string]*Sheet
}

func (ss *Spreadsheet) sheetsByName() map[string]*Sheet {
	m := make(map[string]*Sheet)
	for _, sheet := range ss.Sheets {
		m[sheet.Name] = sheet
	}
	return m
}

// Sheet is a single page in a spreadsheet.
type Sheet struct {
	Name string
	Rows []*Row
}

// Cell returns the j-th cell in the i-th row of sheet s.
func (s *Sheet) Cell(i int, j int) *Cell {
	return s.Rows[i].Cells[j]
}

// Row is a single row of data in a spreadsheet, containing multiple cells.
type Row struct {
	Cells []*Cell
}

// Cell holds a single piece of data in a spreadsheet.
type Cell struct {
	dataType CellDataType
	data     string
}

// CellDataType represents the primitive data type in spreadsheet cells.
type CellDataType uint8

// CellDataType constants.
const (
	String CellDataType = iota
	Numeric
	Bool
)

// Name returns name of cell data type t.
func (t CellDataType) Name() string {
	switch t {
	case String:
		return "String"
	case Numeric:
		return "Numeric"
	case Bool:
		return "Bool"
	default:
		return "Unknown"
	}
}

// Type returns the data type of cell c.
func (c *Cell) Type() CellDataType {
	return c.dataType
}

// Is returns true if data type of cell c is the same with the given type t.
func (c *Cell) Is(t CellDataType) bool {
	return c.dataType == t
}

// String returns the cell data as a string.
func (c *Cell) String() string {
	return c.data
}

// Float returns the cell data as a float.
// If the cell data is not a numeric value, an error is returned.
func (c *Cell) Float() (float64, error) {
	if !c.Is(Numeric) {
		return 0, fmt.Errorf("cell data is not numeric")
	}

	f, err := strconv.ParseFloat(c.data, 64)
	if err != nil {
		return 0, errors.Wrap(err, "failed to convert cell data to number")
	}
	return f, nil
}

// Int returns the cell data as an int.
// If the cell data is not a numeric value, an error is returned.
func (c *Cell) Int() (int, error) {
	if !c.Is(Numeric) {
		return 0, fmt.Errorf("cell data is not numeric")
	}

	f, err := strconv.ParseFloat(c.data, 64)
	if err != nil {
		return 0, errors.Wrap(err, "failed to convert cell data to number")
	}
	return int(f), nil
}

// Int64 returns the cell data as an int64.
// If the cell data is not a numeric value, an error is returned.
func (c *Cell) Int64() (int64, error) {
	if !c.Is(Numeric) {
		return 0, fmt.Errorf("cell data is not numeric")
	}

	i, err := strconv.ParseInt(c.data, 10, 64)
	if err != nil {
		return 0, errors.Wrap(err, "failed to convert cell data to number")
	}
	return i, nil
}

// Bool returns the cell data as a bool.
// If the cell data is not a bool
func (c *Cell) Bool() (bool, error) {
	if !c.Is(Bool) {
		return false, fmt.Errorf("cell data is not bool")
	}

	b, err := strconv.ParseBool(c.data)
	if err != nil {
		return false, errors.Wrap(err, "failed to convert cell data to bool")
	}
	return b, nil
}
