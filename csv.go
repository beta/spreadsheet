// Copyright (c) 2018 Beta Kuang
//
// This software is released under the MIT License.
// https://opensource.org/licenses/MIT

package spreadsheet

import (
	"bufio"
	"encoding/csv"
	"os"

	"github.com/pkg/errors"
)

func openCSVFile(filePath string) (*Spreadsheet, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, errors.Wrap(err, "failed to open file")
	}

	r := csv.NewReader(bufio.NewReader(file))
	records, err := r.ReadAll()
	if err != nil {
		return nil, errors.Wrap(err, "failed to read CSV file")
	}

	s := &Sheet{
		Name: "Sheet 1",
		Rows: make([]*Row, 0, len(records)),
	}
	for _, record := range records {
		row := &Row{
			Cells: make([]*Cell, 0, len(record)),
		}
		for _, data := range record {
			row.Cells = append(row.Cells, &Cell{
				dataType: String,
				data:     data,
			})
		}
		s.Rows = append(s.Rows, row)
	}

	ss := &Spreadsheet{
		Sheets: []*Sheet{s},
	}
	ss.SheetsByName = ss.sheetsByName()
	return ss, nil
}
