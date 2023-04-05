package main

import (
	"fmt"
	"os"

	"github.com/rmasci/csvformat"
)

func main() {
	csvtext := fmt.Sprintf("One,Two,Three\nFour,Five,Six\nSeven,Eight,Nine")
	csvexcel := csvformat.NewExcel()
	csvexcel.Header = "First,Second,Third" // This is optional. If the CSV already contains a heading line. don't set this.
	out, err := csvexcel.Excelout(csvtext, "excelfile.xlsx")
	if err != nil {
		fmt.Println("error:", err)
		os.Exit(1)
	}
	fmt.Println(out)
}
