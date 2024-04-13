package reader

import (
	"fmt"
	"reflect"
	"strconv"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

func ReadFile(filePath string, sheetStruct interface{}) (interface{}, error) {

	xlFile, err := xlsx.OpenFile(filePath)
	if err != nil {
		return err, nil
	}

	//set Sheet by name
	sheet := getSheetsByName(xlFile)

	extractedData := getTemplateType(sheetStruct, sheet, true)

	return extractedData, err
}

func getTemplateType(sheetStruct interface{}, sheet *xlsx.Sheet, skipHeader bool) interface{} {

	// Get the type of struct
	templateType := reflect.TypeOf(sheetStruct)

	// A slice to hold instances of the struct
	templateSliceType := reflect.SliceOf(templateType)
	excelDataSlice := reflect.MakeSlice(templateSliceType, 0, 0)

	// Iterate through rows
	for rI, row := range sheet.Rows {
		// Skip header if
		if skipHeader && rI == 0 {
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
					return err
				}
			}
		}
		excelDataSlice = reflect.Append(excelDataSlice, newStruct)
	}

	// Convert the slice back to interface{}
	return excelDataSlice.Interface()
}

func getSheetsByName(xlFile *xlsx.File) *xlsx.Sheet {
	xf, err := xlFile.Sheet["Sheet1"]

	if err {
		fmt.Println("Couldn't find the sheet name, check the name and run again")
	}

	fmt.Println("Selected Sheet Name:", xf.Name)
	return xf
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
			_ = fmt.Errorf("unsupported type for pointer: %v", field.Type().Elem().Kind())
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
