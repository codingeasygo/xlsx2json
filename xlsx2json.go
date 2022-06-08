package xlsx2json

import (
	"fmt"
	"reflect"
	"strings"
	"time"

	"github.com/tealeg/xlsx"
)

type Field struct {
	Key      string   `json:"key"`
	Required bool     `json:"req"`
	Type     string   `json:"type"`
	Args     []string `json:"args"`
	Row      int      `json:"row"`
	Col      int      `json:"col"`
}

const (
	FieldTypeInt64   = "int64"
	FieldTypeFloat64 = "float64"
	FieldTypeString  = "string"
	FieldTypeFile    = "file"
	FieldTypeTime    = "time"
	FieldTypeRef     = "ref"
)

var FieldTypeAll = []string{FieldTypeInt64, FieldTypeFloat64, FieldTypeString, FieldTypeFile, FieldTypeTime, FieldTypeRef}

type Reader struct {
	File   *xlsx.File
	OnTime func(field *Field, cell *xlsx.Cell) (interface{}, error)
	OnFile func(field *Field, cell *xlsx.Cell) (interface{}, error)
}

func OpenReader(fileName string) (reader *Reader, err error) {
	reader = &Reader{}
	reader.File, err = xlsx.OpenFile(fileName)
	return
}

func (r *Reader) ReadField(sheet string, row int) (fields []*Field, refs []*Field, err error) {
	s := r.File.Sheet[sheet]
	if s == nil {
		err = fmt.Errorf("sheet %v is not exists", sheet)
		return
	}
	cells := s.Rows[row].Cells
	allFields := map[string]*Field{}
	for col, cell := range cells {
		value := strings.TrimSpace(cell.String())
		if len(value) < 1 {
			continue
		}
		parts := strings.Split(value, ",")
		if len(parts) < 2 {
			err = fmt.Errorf("%v is invalid on %v,%v,%v", value, sheet, row, col)
			break
		}
		keyParts := strings.SplitN(parts[0], ":", 2)
		field := &Field{
			Key:      keyParts[0],
			Required: len(keyParts) < 2 || !(strings.ToLower(keyParts[1]) == "o" || strings.ToLower(keyParts[1]) == "optional"),
			Type:     parts[1],
			Args:     parts[2:],
			Row:      row,
			Col:      col,
		}
		if len(field.Key) < 1 {
			err = fmt.Errorf("%v is invalid by key empty on %v,%v,%v", value, sheet, row, col)
			break
		}
		if !strings.Contains("~"+strings.Join(FieldTypeAll, "~")+"~", "~"+field.Type+"~") {
			err = fmt.Errorf("%v is invalid by type not supported on %v,%v,%v", value, sheet, row, col)
			break
		}
		if field.Type == FieldTypeRef && len(field.Args) < 2 {
			err = fmt.Errorf("%v is invalid on %v,%v,%v", value, sheet, row, col)
			break
		}
		if having := allFields[field.Key]; having != nil {
			err = fmt.Errorf("%v is dupcate on %v,%v,%v=>%v,%v,%v", value, sheet, row, col, sheet, having.Row, having.Col)
			break
		}
		allFields[field.Key] = field
		fields = append(fields, field)
		if field.Type == FieldTypeRef {
			refs = append(refs, field)
		}
	}
	return
}

func (r *Reader) objectSet(object map[string]interface{}, field *Field, value interface{}) (err error) {
	parent := object
	keys := strings.Split(field.Key, ".")
	if len(keys) > 1 {
		for i, key := range keys[0 : len(keys)-1] {
			if parent[key] == nil {
				parent[key] = map[string]interface{}{}
			}
			if v, ok := parent[key].(map[string]interface{}); ok {
				parent = v
			} else {
				err = fmt.Errorf("%v is not object", strings.Join(keys[0:i+1], "."))
				break
			}
		}
		if err != nil {
			return
		}
	}
	parent[keys[len(keys)-1]] = value
	return
}

func (r *Reader) readCellRef(s *xlsx.Sheet, field *Field, sheetValues map[string][]map[string]interface{}, object map[string]interface{}, cell *xlsx.Cell) (value interface{}, err error) {
	allValues := []map[string]interface{}{}
	for _, refValue := range sheetValues[field.Args[0]] {
		parts := strings.SplitN(field.Args[1], "=", 2)
		refFieldValue := refValue[strings.TrimSpace(parts[0])]
		if len(parts) > 1 {
			if reflect.DeepEqual(object[parts[1]], refFieldValue) {
				allValues = append(allValues, refValue)
			}
			continue
		}
		switch refFieldValue.(type) {
		case int64:
			var val int64
			val, err = cell.Int64()
			if err != nil {
				break
			}
			if val == refFieldValue {
				allValues = append(allValues, refValue)
			}
		case float64:
			var val float64
			val, err = cell.Float()
			if err != nil {
				break
			}
			if val == refFieldValue {
				allValues = append(allValues, refValue)
			}
		case string:
			if strings.TrimSpace(cell.String()) == refFieldValue {
				allValues = append(allValues, refValue)
			}
		}
	}
	if err == nil {
		value = allValues
	}
	return
}

func (r *Reader) readRowObject(s *xlsx.Sheet, fields []*Field, refs []*Field, sheetValues map[string][]map[string]interface{}, row int) (object map[string]interface{}, err error) {
	object = map[string]interface{}{}
	for _, field := range fields {
		cell := s.Cell(row, field.Col)
		if field.Type == FieldTypeRef {
			object[field.Key], err = r.readCellRef(s, field, sheetValues, object, cell)
			if err != nil {
				break
			}
			continue
		}
		if len(strings.TrimSpace(cell.Value)) < 1 {
			if field.Required {
				err = fmt.Errorf("%v value is empty on %v,%v,%v", field.Key, s.Name, row, field.Col)
				break
			}
			continue
		}
		var value interface{}
		switch field.Type {
		case FieldTypeInt64:
			value, err = cell.Int64()
		case FieldTypeFloat64:
			value, err = cell.Float()
		case FieldTypeString:
			value = strings.TrimSpace(cell.String())
		case FieldTypeTime:
			if r.OnTime != nil {
				value, err = r.OnTime(field, cell)
			} else {
				v, _ := cell.GetTime(false)
				_, zone1 := v.Zone()
				_, zone2 := time.Now().Zone()
				v = v.Add(time.Duration(zone1-zone2) * time.Second).Local()
				value = v
			}
		case FieldTypeFile:
			if r.OnFile != nil {
				value, err = r.OnFile(field, cell)
			} else {
				value = strings.TrimSpace(cell.String())
			}
		}
		if err == nil {
			err = r.objectSet(object, field, value)
		}
		if err != nil {
			break
		}
	}
	return
}

func (r *Reader) readSheet(sheetValues map[string][]map[string]interface{}, sheet string, row, skip int) (data []map[string]interface{}, err error) {
	s := r.File.Sheet[sheet]
	if s == nil {
		err = fmt.Errorf("sheet %v is not exists", sheet)
		return
	}
	fields, refs, err := r.ReadField(sheet, row)
	if err != nil {
		return
	}
	for _, ref := range refs {
		if sheetValues[ref.Args[0]] == nil {
			sheetValues[ref.Args[0]], err = r.Read(ref.Args[0], row, skip)
		}
		if err != nil {
			break
		}
	}
	if err != nil {
		return
	}
	var object map[string]interface{}
	for row := row + skip + 1; ; row++ {
		if len(strings.TrimSpace(s.Cell(row, 0).String())) < 1 { //end
			break
		}
		object, err = r.readRowObject(s, fields, refs, sheetValues, row)
		if err != nil {
			break
		}
		data = append(data, object)
	}
	return
}

func (r *Reader) Read(sheet string, row, skip int) (data []map[string]interface{}, err error) {
	sheetValues := map[string][]map[string]interface{}{}
	data, err = r.readSheet(sheetValues, sheet, row, skip)
	return
}
