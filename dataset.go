package tablib

import (
	"encoding/csv"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

func NewDataSet() *DataSet {
	return &DataSet{
		data: make([][]string, 0, 100),
	}
}

type DataSet struct {
	Title   string
	Headers []string
	data    [][]string
}

func (ds *DataSet) SetTitle(title string) *DataSet {
	ds.Title = title
	return ds
}

func (ds *DataSet) SetHeaders(headers []string) *DataSet {
	ds.Headers = headers
	return ds
}

func (ds *DataSet) Append(row []string) (err error) {
	if len(ds.Headers) > 0 && len(row) != len(ds.Headers) {
		err = ErrInvalidRow
		return
	}
	ds.data = append(ds.data, row)
	return
}

func (ds *DataSet) AppendCol(col []string, header string) (err error) {
	if len(col) != len(ds.data) {
		err = ErrInvalidCol
		return
	}
	if len(ds.Headers) > 0 {
		ds.Headers = append(ds.Headers, header)
	}
	for i, _ := range ds.data {
		ds.data[i] = append(ds.data[i], col[i])
	}
	return
}

func (ds *DataSet) Records() [][]string {
	return ds.data
}

func (ds *DataSet) Len() int {
	return len(ds.data)
}

func (ds *DataSet) GetColByIndex(idx int) []string {
	if idx < 0 || idx >= len(ds.data[0]) {
		return nil
	}
	col := make([]string, 0, len(ds.data))
	for _, row := range ds.data {
		col = append(col, row[idx])
	}
	return col
}

func (ds *DataSet) GetColByHeader(header string) []string {
	if len(ds.Headers) == 0 {
		return nil
	}
	idx := -1
	for i, h := range ds.Headers {
		if h == header {
			idx = i
			break
		}
	}
	if idx == -1 {
		return nil
	}
	return ds.GetColByIndex(idx)
}

func (ds *DataSet) Load(input io.Reader, format Format, withHeaders bool) (*DataSet, error) {
	var err error
	switch format {
	case CSV:
		err = ds.loadCSV(input, withHeaders)
	case XLSX:
		err = ds.loadXLSX(input, withHeaders)
	default:
		err = ErrUnsupportedFormat
	}
	return ds, err
}

func (ds *DataSet) Export(output io.Writer, format Format) (err error) {
	switch format {
	case CSV:
		err = ds.exportCSV(output)
	case XLSX:
		err = ds.exportXLSX(output)
	default:
		err = ErrUnsupportedFormat
		return
	}
	return
}

func (ds *DataSet) loadCSV(input io.Reader, withHeaders bool) (err error) {
	records, err := csv.NewReader(input).ReadAll()
	if err != nil {
		return
	}
	ds.loadData(records, withHeaders)
	return
}

func (ds *DataSet) loadXLSX(input io.Reader, withHeaders bool) (err error) {
	f, err := excelize.OpenReader(input)
	if err != nil {
		return
	}
	defer f.Close()
	ds.Title = f.GetSheetList()[0]
	rows, err := f.GetRows(f.GetSheetList()[0])
	if err != nil {
		return
	}
	ds.loadData(rows, withHeaders)
	return
}

func (ds *DataSet) loadData(rows [][]string, withHeaders bool) {
	if len(rows) == 0 {
		return
	}
	if withHeaders {
		ds.Headers = rows[0]
		if len(rows) > 1 {
			ds.data = rows[1:]
		}
	} else {
		ds.data = rows
	}
	return
}

func (ds *DataSet) exportCSV(output io.Writer) (err error) {
	writer := csv.NewWriter(output)
	if len(ds.Headers) > 0 {
		err = writer.Write(ds.Headers)
	}
	if err != nil {
		return
	}
	err = writer.WriteAll(ds.data)
	if err != nil {
		return
	}
	return
}

func (ds *DataSet) exportXLSX(output io.Writer) (err error) {
	f := excelize.NewFile()
	defer f.Close()
	if ds.Title == "" {
		ds.Title = "Sheet1"
	}
	f.SetSheetName("Sheet1", ds.Title)
	streamWriter, err := f.NewStreamWriter(ds.Title)
	if err != nil {
		return
	}
	err = ds.exportStreamWriter(streamWriter)
	if err != nil {
		return
	}
	err = f.Write(output)
	return
}

func (ds *DataSet) exportStreamWriter(sw *excelize.StreamWriter) (err error) {
	rowStart := 1
	if len(ds.Headers) > 0 {
		_headers := make([]interface{}, 0, len(ds.Headers))
		for _, h := range ds.Headers {
			_headers = append(_headers, excelize.Cell{Value: h})
		}
		err = sw.SetRow(fmt.Sprintf("A%d", rowStart), _headers)
		if err != nil {
			return
		}
		rowStart++
	}
	for i, row := range ds.data {
		_row := make([]interface{}, 0, len(row))
		for _, r := range row {
			_row = append(_row, excelize.Cell{Value: r})
		}
		err1 := sw.SetRow(fmt.Sprintf("A%d", i+rowStart), _row)
		if err1 != nil {
			err = err1
			return
		}
	}
	err = sw.Flush()
	return
}
