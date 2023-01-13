# csvformat

One thing that takes up a ton of time is getting the output to look right. With csvformat you can take any text string that is formatted in CSV and create a nice table format.
Example: 
```cgo
func main() {
	csvstuff := getCsv() // This could be any function that returns a text string with CSV.
	csvout := csvformat.NewGrid()
	fmt.Println(csvout.Gridout(string(csvstuff)))
}
```
See minimal.go in the ex folder.
Function getCsv() just does a simple http.Client to https://cdn.wsform.com/wp-content/uploads/2021/04/weekday.csv
The main point is that you just need to format it in csv, and you don't have to do any additional formatting.

Output:
```cgo
+--------------+-----------------+------------+--------------+
|     Name     |   Abbreviation  |   Numeric  |   Numeric-2  |
+==============+=================+============+==============+
|    Monday    |       Mon.      |      1     |      01      |
|    Tuesday   |       Tue.      |      2     |      02      |
|   Wednesday  |       Wed.      |      3     |      03      |
|   Thursday   |       Thu.      |      4     |      04      |
|    Friday    |       Fri.      |      5     |      05      |
|   Saturday   |       Sat.      |      6     |      06      |
|    Sunday    |       Sun.      |      7     |      07      |
+--------------+-----------------+------------+--------------+

```
Better Example:
``` go 
func main() {
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
```
Output:
```cgo
+----------+-----------+----------+
|   First  |   Second  |   Third  |
+==========+===========+==========+
|    One   |    Two    |   Three  |
|   Four   |    Five   |    Six   |
|   Seven  |   Eight   |   Nine   |
+----------+-----------+----------+

```

Output formats:
* simple
* plain
* tab
* text
* html Todo: format as a bootstrap table.
* mysql
