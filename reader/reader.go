package structurama

import (
	"fmt"
	"github.com/tealeg/xlsx"
	"reflect"
	"strconv"
	"strings"
	"time"
)

var sheet *xlsx.Sheet
var sheetStruct interface{}
var _skipHeader bool

func ReadFileDefault(filePath string, customStruct interface{}, skipHeader bool, sheetNumber ...int) (interface{}, error) {
	sheetStruct = customStruct
	_skipHeader = skipHeader
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	if len(sheetNumber) > 0 {
		sheet, err = getSheetByNumber(xlFile, sheetNumber[0])
		if err != nil {
			return nil, err
		}
	} else {
		sheet = xlFile.Sheets[0]
	}

	data, err := read(sheet)
	if err != nil {
		return nil, err
	}

	return data, nil
}

func ReadFileBySheetName(filePath string, customStruct interface{}, skipHeader bool, sheetName string) (interface{}, error) {
	sheetStruct = customStruct
	_skipHeader = skipHeader
	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		return nil, err
	}

	sheet, err = getSheetByName(xlFile, sheetName)
	if err != nil {
		return nil, err
	}

	data, err := read(sheet)
	if err != nil {
		return nil, err
	}

	return data, nil
}

func read(sheet *xlsx.Sheet) (interface{}, error) {
	// Get the type of struct
	templateType := reflect.TypeOf(sheetStruct)

	// A slice to hold instances of the struct
	templateSliceType := reflect.SliceOf(templateType)
	excelDataSlice := reflect.MakeSlice(templateSliceType, 0, 0)

	// Iterate through rows
	for r_i, row := range sheet.Rows {
		// Skip header if
		if _skipHeader && r_i == 0 {
			continue
		}

		// Create a new instance of the struct
		newStruct := reflect.New(templateType).Elem()

		for i, cell := range row.Cells {
			// Get the field corresponding to the current column index
			field := newStruct.FieldByIndex([]int{i})
			if field.IsValid() {
				err := handleFieldKind(field, cell)
				if err != nil {
					return nil, err
				}
			}
		}
		excelDataSlice = reflect.Append(excelDataSlice, newStruct)
	}

	// Convert the slice back to interface{}
	return excelDataSlice.Interface(), nil
}

func isValidSheetNumber(sheetNumber int, totalSheets int) bool {
	return sheetNumber >= 0 && sheetNumber < totalSheets
}

func getSheetByNumber(xlFile *xlsx.File, sheetNumber int) (*xlsx.Sheet, error) {
	if isValidSheetNumber(sheetNumber, len(xlFile.Sheets)) {
		return xlFile.Sheets[sheetNumber], nil
	}
	return nil, fmt.Errorf("invalid sheet number: %v", sheetNumber)
}

func getSheetByName(xlFile *xlsx.File, sheetName string) (*xlsx.Sheet, error) {
	if xlFile.Sheet[sheetName] != nil {
		return xlFile.Sheet[sheetName], nil
	}
	return nil, fmt.Errorf("invalid sheet name: %s", sheetName)
}

// handleFieldKind: Convert cell value based on the field type
func handleFieldKind(field reflect.Value, cell *xlsx.Cell) error {
	switch field.Kind() {
	case reflect.String:
		field.SetString(cell.String())
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		intValue, err := strconv.ParseInt(cell.String(), 10, 64)
		if err == nil {
			field.SetInt(intValue)
		}
	case reflect.Float32, reflect.Float64:
		floatValue, err := strconv.ParseFloat(cell.String(), 64)
		if err == nil {
			field.SetFloat(floatValue)
		}
	case reflect.Ptr:
		// Handle pointer types more explicitly
		if field.Type().Elem().Kind() == reflect.String {
			ptrValue := reflect.New(field.Type().Elem()).Elem()
			ptrValue.SetString(cell.String())
			field.Set(ptrValue.Addr())
		} else if field.Type().Elem().Kind() == reflect.Int {
			ptrValue := reflect.New(field.Type().Elem()).Elem()
			if cell.String() != "" {
				intValue, err := strconv.ParseInt(cell.String(), 10, 64)
				if err == nil {
					ptrValue.SetInt(intValue)
					field.Set(ptrValue.Addr())
				}
			}
		} else {
			return fmt.Errorf("unsupported type for pointer: %v", field.Type().Elem().Kind())
		}

	case reflect.Struct:
		if field.Type() == reflect.TypeOf(time.Time{}) {
			if strings.Contains(field.String(), "-") {
				dateValue, err := time.Parse("2006-01-02", cell.String())
				if err == nil {
					field.Set(reflect.ValueOf(dateValue))
				}
			} else if strings.Contains(field.String(), "/") {
				dateValue, err := time.Parse("2006/01/02", cell.String())
				if err == nil {
					field.Set(reflect.ValueOf(dateValue))
				}
			} else {
				dateValue, err := time.Parse("02012006", cell.String())
				if err == nil {
					field.Set(reflect.ValueOf(dateValue))
				}

			}
		}
	default:
		field.SetString(cell.String())
	}
	return nil
}
