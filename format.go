package csvformat

import (
	"bufio"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/rmasci/excelize/v2"
	//"github.com/rmasci/excelize/v2"
	"github.com/rmasci/gotabulate"
)

// Grid Render Used for Gridout function.  Renders output text in a grid, text, tab.
//
//		Render:
//			simple Simple tab delimited with underlined header.
//			plain pretty much just strips the "," for a space
//			tab   Tab spaced output
//			html  Output in simple HTML. Header has grey background with thin lines between cells
//			mysql Looks similar to mysql shell output
//			grid  uses ANSI Graphics. Not compatible with all terminals
//			gridt MySQL
//		Align: Right, left or center
//		NoHeader: don't print a header
//		Wrap: wrap cell output
//		NoLineBetweenRow: put a blank line inbetween each row
//		Columns: List the columns you want output.
//		Space: Do not trim spaces from column. "1,   2" will output "1|    2" instead of "1|2"
//		Delimeter: Delimeter between text. Default is a ","
//	 Number -- automatically add number the output
type Grid struct {
	Render           string
	Align            string
	NoHeader         bool
	Wrap             bool
	Indent           string
	NoLineBetweenRow bool
	Space            bool
	Columns          string
	Delimiter        string
	Number           bool
	OutDelimiter     string
}

type Excel struct {
	Header    string
	FileName  string
	Space     bool
	Delimeter string
	Sheet     string
}

// TimeStamp
// Returns the current time as timestamp: %Y%m%d%H%M%s  optionally you can give it a time format in accordance with Go's time package.
func TimeStamp(a ...any) string {
	var f string
	if len(a) >= 1 {
		f = fmt.Sprintf("%v", a[0])
	} else {
		f = "20060102150405"
	}
	return time.Now().Format(f)
}

// NewGrid Returns a Grid with some defaults.
//
//	Render: mysql
//	Align: left
//	NoHeader: False
//	Wrap: false
//	Indent: "" (empty)
//	LineBetweenRow: false
//	Columns: All (empty)
//
// New grid gives defaults. You can change these after NewGrid or just set it using the struct as shown in NewGrid.
// your own use might look like this
// mygrid := nsres.Grid{Render: "grid", Align: "Right", NoHeader: false, Wrap: false, Indent: "\t", LineBetweenRow: false, Space: false, Columns: "", Delimiter: "|"
func NewGrid() *Grid {
	g := Grid{Render: "mysql", Align: "Left", NoHeader: false, Wrap: false, Indent: "", NoLineBetweenRow: true, Space: true, Columns: "", Delimiter: ","}
	return &g
}

// NewExcel Returns type Excel with some defaults.
//
//	Align: left
//	NoHeader: False
//	Wrap: false
//	Indent: "" (empty)
//	LineBetweenRow: false
//	Columns: All (empty)
//
// New Excel gives defaults. You can change these after NewExcel or just set it using the struct as shown in NewExcel.
// your own use might look like this
// myexcek := nsres.Excel{Delimiter: "|"
func NewExcel() *Excel {
	e := Excel{FileName: "mysql", Header: "", Space: true, Delimeter: ",", Sheet: "Sheet1"}
	return &e
}

// Gridout
// Don't waste time getting the output looking nice, just make it csv, then pass it to gridoutl.  Gridout Prints according to Grid struct.
// Render:
// simple. Tab format with lines on top
// plain. Text output no gridlines extra spaces in between
//
func (g *Grid) Gridout(text string) (string, error) {
	var out string
	//if g.Render == "text" {
	//	//g.Render = "csv"
	//	g.Render = "txt"
	//}
	txtRdr := strings.NewReader(text)
	csvRdr := csv.NewReader(txtRdr)
	csvRdr.Comma = []rune(g.Delimiter)[0]
	csvRdr.TrimLeadingSpace = g.Space
	s, err := csvRdr.ReadAll()
	if len(s) <= 1 {
		g.NoHeader = true
	}
	if err != nil {
		return "", fmt.Errorf("Error csvRdr.ReadAll: %v\n", err)
	}
	// Normalize for errors since we get an error trying to convert a non number to string.
	g.Columns = strings.TrimSpace(g.Columns)
	for i, _ := range s {
		if g.Columns != "" {
			var tmp []string
			pCol := strings.Split(g.Columns, ",")
			for _, c := range pCol {
				c = strings.TrimSpace(c)
				idx, err := strconv.Atoi(c)
				if err != nil {
					//return "", fmt.Errorf("can't convert %v to int %err", c, err)
					continue
				}
				if idx > len(s[i]) || idx <= 0 {
					err := fmt.Errorf("invalid column col:%v len:%v", idx, len(s[i]))
					return "", err
				}
				idx = idx - 1
				if g.Render == "bingo" {
					s[i][idx] = fmt.Sprintf("\n\n%s\n\n", s[i][idx])
				}
				tmp = append(tmp, s[i][idx])
			}
			s[i] = tmp
		}
		if i == 0 && g.Number {
			s[i] = append([]string{"Index"}, s[i]...)
		} else if g.Number {
			I := fmt.Sprintf("%d", i)
			s[i] = append([]string{I}, s[i]...)
		}
	}
	gridulate := gotabulate.Create(s)
	gridulate.SetAlign(g.Align)
	gridulate.SetWrapStrings(g.Wrap)
	gridulate.SetRemEmptyLines(g.NoLineBetweenRow)
	gridulate.NoHeader = g.NoHeader
	scanner := bufio.NewScanner(strings.NewReader(gridulate.Render(g.Render)))
	for scanner.Scan() {
		out = fmt.Sprintf("%s%s%s\n", out, g.Indent, scanner.Text())
	}
	return out, nil
}

func (excel *Excel) Excelout(text string, fname string) (string, error) {
	if excel.Header != "" {
		text = fmt.Sprintf("%s\n%s", excel.Header, text)
	}
	lines := strings.Split(strings.TrimSpace(text), "\n")
	fmt.Printf("Len: %v\n", len(lines))
	if excel.FileName == "" {
		return "", fmt.Errorf("must pass a file name to store the Excel in.\n")
	}
	xlsx := excelize.NewFile()
	xlsx.NewSheet(excel.Sheet)
	lineNumber := 0
	/// This is each row
	for _, row := range lines {
		for i, cell := range strings.Split(row, excel.Delimeter) {
			// each cell in row
			if excel.Space {
				cell = strings.TrimSpace(cell)
			}
			cntn, err := excelize.ColumnNumberToName(i + 1) //ColumnNumberToName(i + 1)
			if err != nil {
				continue
			}
			axis := fmt.Sprintf("%v%v", cntn, lineNumber)
			xlsx.SetCellValue(excel.Sheet, axis, cell)
		}
		lineNumber++
	}
	if err := xlsx.SaveAs(fname); err != nil {
		return "", err
	}
	return fmt.Sprintf("Wrote %v rows to %v\n", lineNumber, fname), nil
}

func Jsonout(text string, pp bool) (string, error) {
	rows := strings.Split(text, "\n")
	c1 := strings.Split(rows[0], ",")
	rows = rows[1:]
	tableData := make([]map[string]interface{}, 0)
	//values := make([]interface{}, count)
	//valuePtrs := make([]interface{}, count)
	for _, line := range rows {
		entry := make(map[string]interface{})
		cells := strings.Split(line, ",")
		for i, col := range cells {
			strings.TrimSpace(col)
			entry[c1[i]] = col
		}
		tableData = append(tableData, entry)
	}
	if len(tableData) <= 1 {
		fmt.Printf("0 Rows...")
		os.Exit(0)
	}
	if pp {
		jsonData, err := json.MarshalIndent(tableData, "", "\t")
		if err != nil {
			return "", err
		}
		return fmt.Sprintln(string(jsonData)), nil
	} else {
		jsonData, err := json.Marshal(tableData)
		if err != nil {
			return "", err
		}
		return fmt.Sprintln(string(jsonData)), nil
	}
	return "", nil
}
