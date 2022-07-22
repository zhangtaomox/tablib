package tablib

import "errors"

var ErrUnsupportedFormat = errors.New("unsupported format")
var ErrInvalidRow = errors.New("invalid row")
var ErrInvalidCol = errors.New("invalid col")
