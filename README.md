# Spreadsheet

Spreadsheet is a Go package providing a simple interface for reading spreadsheet files, including XLSX and CSV.

## Installation

Please use [gopkg.in](https://gopkg.in) for stable releases.

```bash
go get -u gopkg.in/beta/spreadsheet.v1
```

## Getting started

```go
package main

import (
	"fmt"

	"gopkg.in/beta/spreadsheet.v1"
)

func main() {
	// Open sheet.
	ss, err := spreadsheet.Open("test.xls")
	if err != nil {
		...
	}

	s1 := ss.Sheets[0]
	fmt.Printf("Sheet name: %s", s1.Name)
	// Traverse sheet.
	for _, row := range s1.Rows {
		for _, cell := range row.Cells {
			switch cell.Type() {
			case spreadsheet.String:
				fmt.Printf("Value: %s\n", cell.String())
			case spreadsheet.Numeric:
				f, _ := cell.Float()
				fmt.Printf("Value: %f\n", f)
				i, _ := cell.Int()
				fmt.Printf("Value: %d\n", i)
				i64, _ := cell.Int64()
				fmt.Printf("Value: %d\n", i64)
			case spreadsheet.Bool:
				b, _ := cell.Bool()
				fmt.Printf("Value: %t\n", b)
			}
		}
	}

	s2 := ss.SheetsByName["Sheet 2"] // Pick sheet by name.
	cell := s2.Cell(2, 3)            // Pick cell by coordinate.
	if cell.Is(spreadsheet.String) { // Check cell type.
		fmt.Printf("Value: %s", cell.String())
	}
}
```

## License

MIT