package xlsx2json

import (
	"encoding/json"
	"fmt"
	"testing"

	"github.com/tealeg/xlsx"
)

func JSON(v interface{}) string {
	data, err := json.Marshal(v)
	if err != nil {
		return err.Error()
	}
	return string(data)
}

func TestReader(t *testing.T) {
	reader, err := OpenReader("example.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	{ //id ref
		users, err := reader.Read("user0", 1, 0)
		if err != nil {
			t.Error(err)
			return
		}
		for _, user := range users {
			if user["products"] == nil {
				t.Error("error")
				return
			}
			if v, ok := user["products"].([]map[string]interface{}); !ok || len(v) < 1 {
				t.Error("error")
				return
			}
		}
		fmt.Println(JSON(users))
	}
	{ //string ref
		data, err := reader.Read("user1", 1, 0)
		if err != nil {
			t.Error(err)
			return
		}
		for _, item := range data {
			if item["products_0"] == nil || item["products_1"] == nil || item["products_2"] == nil {
				t.Error("error")
				return
			}
			if v, ok := item["products_0"].([]map[string]interface{}); !ok || len(v) < 1 {
				t.Error("error")
				return
			}
			if v, ok := item["products_1"].([]map[string]interface{}); !ok || len(v) < 1 {
				t.Error("error")
				return
			}
			if v, ok := item["products_2"].([]map[string]interface{}); !ok || len(v) < 1 {
				t.Error("error")
				return
			}
		}
		fmt.Println(JSON(data))
	}
	{
		reader.OnFile = func(field *Field, cell *xlsx.Cell) (interface{}, error) {
			return cell.String(), nil
		}
		reader.OnTime = func(field *Field, cell *xlsx.Cell) (interface{}, error) {
			return cell.GetTime(false)
		}
		data, err := reader.Read("user0", 1, 0)
		if err != nil {
			t.Error(err)
			return
		}
		fmt.Println(JSON(data))
	}
	{
		_, err = reader.Read("user2", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
		reader.OnParse = func(field *Field, cell *xlsx.Cell) (interface{}, error) {
			return cell.String(), nil
		}
		data, err := reader.Read("user2", 1, 0)
		if err != nil {
			t.Error(err)
			return
		}
		fmt.Println(JSON(data))
	}
}

func TestError(t *testing.T) {
	reader, err := OpenReader("test_err_0.xlsx")
	if err != nil {
		t.Error(err)
		return
	}
	//read field
	{
		//
		_, _, err = reader.ReadField("none", 1)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, _, err = reader.ReadField("field0", 1)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, _, err = reader.ReadField("field2", 1)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, _, err = reader.ReadField("field3", 1)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, _, err = reader.ReadField("field4", 1)
		if err == nil {
			t.Error(err)
			return
		}
	}
	//read value
	{
		//
		_, err = reader.Read("value0", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, err = reader.Read("value1", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, err = reader.Read("value2", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
	}
	//read ref
	{
		//
		_, err = reader.Read("ref0", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, err = reader.Read("ref1", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
		//
		_, err = reader.Read("ref2", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
	}
	//read sheet
	{
		//
		_, err = reader.Read("none", 1, 0)
		if err == nil {
			t.Error(err)
			return
		}
	}
}
