package main

import (
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"

	"github.com/rmasci/csvformat"
)

func main() {
	csvstuff := getCsv() // This could be any function that returns a text string with CSV.
	csvout := csvformat.NewGrid()
	fmt.Println(csvout.Gridout(string(csvstuff)))
}

func getCsv() string {
	earl, err := url.Parse("https://cdn.wsform.com/wp-content/uploads/2021/04/weekday.csv")
	errHandle(err, "url parse", true)
	req := http.Request{URL: earl}
	resp, err := http.DefaultClient.Do(&req) // handle the error
	errHandle(err, "http.DefaultClient.Do", true)
	bod, err := io.ReadAll(resp.Body)
	errHandle(err, "io.ReadAll", true)
	return string(bod)
}

// I put this in all my code.
func errHandle(err error, str string, ex bool) bool {
	if err != nil {
		fmt.Println("error", err)
		if ex {
			os.Exit(1)
			return true
		}
	}
	return false
}
