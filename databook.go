package tablib

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"io"
)

func NewDataBook() *DataBook {
	return &DataBook{
		sheets: make([]*DataSet, 0, 4),
	}
}

type DataBook struct {
	sheets []*DataSet
}

func (db *DataBook) AddSheet(dataset *DataSet) {
	db.sheets = append(db.sheets, dataset)
	return
}

func (db *DataBook) Sheets() []*DataSet {
	return db.sheets
}

func (db *DataBook) Load(input io.Reader, format Format, withHeaders bool) (*DataBook, error) {
	var err error
	switch format {
	case XLSX:
		err = db.loadXLSX(input, withHeaders)
	default:
		err = ErrUnsupportedFormat
	}
	return db, err
}

func (db *DataBook) loadXLSX(input io.Reader, withHeaders bool) (err error) {
	f, err := excelize.OpenReader(input)
	if err != nil {
		return
	}
	defer f.Close()
	for _, name := range f.GetSheetList() {
		ds := NewDataSet().SetTitle(name)
		rows, err1 := f.GetRows(name)
		if err1 != nil {
			err = err1
			return
		}
		ds.loadData(rows, withHeaders)
		db.AddSheet(ds)
	}
	return
}

func (db *DataBook) Export(output io.Writer, format Format) (err error) {
	switch format {
	case XLSX:
		err = db.exportXLSX(output)
	default:
		err = ErrUnsupportedFormat
		return
	}
	return
}

func (db *DataBook) exportXLSX(output io.Writer) (err error) {
	f := excelize.NewFile()
	defer f.Close()
	for i, ds := range db.sheets {
		if ds.Title == "" {
			ds.Title = fmt.Sprintf("Sheet%d", i+1)
		}
		if i == 0 {
			f.SetSheetName("Sheet1", ds.Title)
		} else {
			f.NewSheet(ds.Title)
		}
		streamWriter, err1 := f.NewStreamWriter(ds.Title)
		if err1 != nil {
			err = err1
			return
		}
		err1 = ds.exportStreamWriter(streamWriter)
		if err1 != nil {
			err = err1
			return
		}
	}
	err = f.Write(output)
	return
}
