package main

import (
	"fmt"
	"os"

	"github.com/rmasci/csvformat"
)

func main() {
	rendr := "mysql" // Give a mysql shell like table output.
	if len(os.Args) > 1 {
		rendr = os.Args[1] // you can experiment with the output formats
	}
	csvtext := fmt.Sprintf("One,Two,Three\nFour,Five,Six\nSeven,Eight,Nine")
	csvgrid := csvformat.NewGrid()
	csvgrid.Headline = "First,Second,Third" // This is optional. If the CSV already contains a heading line. don't set this.
	csvgrid.Render = rendr                  // There are several different formats.
	out, err := csvgrid.Gridout(csvtext)
	if err != nil {
		fmt.Println("error:", err)
		os.Exit(1)
	}
	fmt.Println(out)
}
