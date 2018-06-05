// Copyright (c) 2018 Beta Kuang
//
// This software is released under the MIT License.
// https://opensource.org/licenses/MIT

package spreadsheet

import (
	"github.com/pkg/errors"
	"github.com/tealeg/xlsx"
)

func openXLSXFile(filePath string) (*Spreadsheet, error) {
	file, err := xlsx.OpenFile(filePath)
	if err != nil {
		return nil, errors.Wrap(err, "failed to open XLSX file")
	}

	ss := &Spreadsheet{
		Sheets: make([]*Sheet, 0, len(file.Sheets)),
	}

	for _, sheet := range file.Sheets {
		s := &Sheet{
			Name: sheet.Name,
			Rows: make([]*Row, 0, len(sheet.Rows)),
		}
		for _, row := range sheet.Rows {
			r := &Row{
				Cells: make([]*Cell, 0, len(row.Cells)),
			}
			for _, cell := range row.Cells {
				c := new(Cell)
				switch cell.Type() {
				case xlsx.CellTypeNumeric:
					c.dataType = Numeric
				case xlsx.CellTypeBool:
					c.dataType = Bool
				default:
					c.dataType = String
				}
				c.data = cell.Value
				r.Cells = append(r.Cells, c)
			}
			s.Rows = append(s.Rows, r)
		}
		ss.Sheets = append(ss.Sheets, s)
	}

	ss.SheetsByName = ss.sheetsByName()
	return ss, nil
}
