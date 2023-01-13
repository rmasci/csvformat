package main

import (
	"fmt"
	"os"

	"github.com/rmasci/csvformat"
)

func main() {
	renders := []string{"simple", "plain", "tab", "text", "csv", "html", "mysql", "mysqlg", "grid", "bingo", "gridt", "simple-nohead"}
	for _, r := range renders {
		fmt.Println("Render:", r)
		g := csvformat.NewGrid()
		g.Render = r
		//g.NoHeader = true
		out, err := g.Gridout(`one,two,three
four,five,six`)
		if err != nil {
			fmt.Println("error", err)
			os.Exit(1)
		}
		fmt.Println(out)
	}
}
