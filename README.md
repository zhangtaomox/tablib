# tablib
go pkg for Tabular Datasets in XLSX, CSV. Salute https://github.com/jazzband/tablib

## Installation
```bash
go get github.com/zhangtaomox/tablib
```
## Quick Start
### Create a new dataset
```go
package main
func main() {
    dataset := tablib.NewDataSet().SetTitle("tablib").SetHeaders([]string{"a", "b", "c"})
}
```
### Add a row
```go
// should not ignore in production code
_ = dataset.Append([]string{"1", "2", "3"})
```
### Add a column
```go
// should not ignore in production code
// note that len(col) must equal to len(ds.Len())
_ = dataset.AppendCol([]string{"4"}, "d")
```
### Get Data
```go
fmt.Println(dataset.Headers)
// [a b c d]
fmt.Println(dataset.Len())
// 1
fmt.Println(dataset.Records())
// [[1 2 3 4]]
```
## Export
```go
// xlsx
fXlsx, _ := os.OpenFile("test.xlsx", os.O_CREATE|os.O_WRONLY, 0666)
defer fXlsx.Close()
_ = dataset.Export(fXlsx, tablib.XLSX)
// csv
fCsv, _ = os.OpenFile("test.csv", os.O_CREATE|os.O_WRONLY, 0666)
defer fCsv.Close()
_ = dataset.Export(fCsv, tablib.CSV)
```

## Load from exists file
```go
// xlsx
dataset2, _ := tablib.NewDataSet().Load(fXlsx, tablib.XLSX, true)
fmt.Println(dataset2.Headers)
// [a b c d]

//csv ...
```
## Databook
databook only support xlsx

### create and export
```go
databook := tablib.NewDataBook()
databook.AddSheet(dataset2)
fmt.Println(databook.Sheets[0].Headers)
// [a b c d]
_ = databook.Export(fDatabook, tablib.XLSX)
```
