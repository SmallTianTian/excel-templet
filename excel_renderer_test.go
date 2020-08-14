package xlsxt

import (
	"bytes"
	"context"
	"fmt"
	"reflect"
	"testing"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize"
)

func writeExcelHelper(data [][]string) ([]byte, error) {
	f := excelize.NewFile()
	for j, item := range data {
		for k, it := range item {
			n, _ := excelize.CoordinatesToCellName(k+1, j+1)
			f.SetCellValue("Sheet1", n, it)
		}
	}
	if bf, err := f.WriteToBuffer(); err != nil {
		return nil, err
	} else {
		return bf.Bytes(), nil
	}
}

func checkExcelHelper(data []byte, expect [][]string) error {
	f, err := excelize.OpenReader(bytes.NewBuffer(data))
	if err != nil {
		return err
	}
	have := make([][]string, 0)
	row, _ := f.Rows("Sheet1")
	for row.Next() {
		s, _ := row.Columns()
		have = append(have, s)
	}
	if len(have) != len(expect) {
		return fmt.Errorf("Line not equal. Except %d != %d.", len(expect), len(have))
	}
	for i, item := range expect {
		if len(item) == 0 && len(have[i]) == 0 {
			continue
		}
		if !reflect.DeepEqual(item, have[i]) {
			return fmt.Errorf("In %d line. Except %v != %v", i, item, have[i])
		}
	}
	return nil
}

func map2InterChanHelper(data []interface{}) chan interface{} {
	c := make(chan interface{}, len(data))
	for _, item := range data {
		c <- item
	}
	close(c)
	return c
}

func map2MapChanHelper(data []map[string]interface{}) chan map[string]interface{} {
	c := make(chan map[string]interface{}, len(data))
	for _, item := range data {
		c <- item
	}
	close(c)
	return c
}

func Test_Rander(t *testing.T) {
	deadlinectx, c := context.WithDeadline(context.Background(), time.Now().Add(-time.Hour))
	c()
	type args struct {
		ctx  context.Context
		temp [][]string
		data interface{}
	}
	tests := []struct {
		name    string
		args    args
		wantRes [][]string
		wantErr bool
	}{
		{
			name: "Base",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"s": "s1", "d": "d1"},
						{"s": "s2", "d": "d2"},
						{"s": "s3", "d": "d3"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "data not map[string]interface{}",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string][]map[string]interface{}{
					"rows": {
						{"s": "s1", "d": "d1"},
						{"s": "s2", "d": "d2"},
						{"s": "s3", "d": "d3"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "template with empty line",
			args: args{
				temp: [][]string{
					{},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string][]map[string]interface{}{
					"rows": {
						{"s": "s1", "d": "d1"},
						{"s": "s2", "d": "d2"},
						{"s": "s3", "d": "d3"},
					},
				},
			},
			wantRes: [][]string{
				{},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "template with no range end",
			args: args{
				temp: [][]string{
					{},
					{"{{range rows}}"},
					{"string", `{{s}}`},
				},
				data: map[string][]map[string]interface{}{
					"rows": {
						{"s": "s1", "d": "d1"},
					},
				},
			},
			wantErr: true,
		},
		{
			name: "template with blank range",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"{{end}}"},
				},
				data: map[string][]map[string]interface{}{},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "template with blank range and result is blank excel",
			args: args{
				temp: [][]string{
					{"{{range rows}}"},
					{"{{end}}"},
				},
				data: map[string][]map[string]interface{}{},
			},
			wantRes: [][]string{},
		},
		{
			name: "data without range key",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range nokey}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "more range",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"split"},
					{"{{range newrows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string][]map[string]interface{}{
						{"newrows": {
							{"s": "s1", "d": "d1"},
							{"s": "s2", "d": "d2"},
							{"s": "s3", "d": "d3"}},
						},
						{"newrows": {
							{"s": "s4", "d": "d4"},
							{"s": "s5", "d": "d5"},
							{"s": "s6", "d": "d6"}},
						},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"split"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
				{"split"},
				{"string", "s4"},
				{"string", "s5"},
				{"string", "s6"},
			},
		},
		{
			name: "range next line is blank",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"s": "s1", "d": "d1"},
						{"s": "s2", "d": "d2"},
						{"s": "s3", "d": "d3"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{},
				{"string", "s1"},
				{},
				{"string", "s2"},
				{},
				{"string", "s3"},
			},
		},
		{
			name: "data with sheet name",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"Sheet1": map[string]interface{}{
						"rows": []map[string]interface{}{
							{"s": "s1", "d": "d1"},
							{"s": "s2", "d": "d2"},
							{"s": "s3", "d": "d3"},
						},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "data is nil",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: nil,
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "range data item not map",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []string{
						"this is string",
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "range data not array or slice or chan",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": "this is string",
				},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "range data item not string key map",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[int]string{
						{1: "a"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "data not string key map",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[int]interface{}{
					1: []map[int]string{
						{1: "a"},
					},
				},
			},
			wantErr: true,
		},
		{
			name: "ctx with de",
			args: args{
				ctx: deadlinectx,
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: nil,
			},
			wantErr: true,
		},
		{
			name: "render chan map[string]interface{}",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": map2MapChanHelper([]map[string]interface{}{
						{"s": "s1", "d": "d1"},
						{"s": "s2", "d": "d2"},
						{"s": "s3", "d": "d3"},
					}),
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "render chan interface{}",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": map2InterChanHelper([]interface{}{
						map[string]string{"s": "s1", "d": "d1"},
						map[string]string{"s": "s2", "d": "d2"},
						map[string]string{"s": "s3", "d": "d3"},
					}),
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
			},
		},
		{
			name: "render chan interface{} with interface{} not map",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": map2InterChanHelper([]interface{}{
						"a",
					}),
				},
			},
			wantRes: [][]string{
				{"Test"},
			},
		},
		{
			name: "sheet data not map",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"Sheet1": "a",
				},
			},
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			bs, err := writeExcelHelper(tt.args.temp)
			if err != nil {
				t.Error(err)
			}
			xl, err := NewFromBinary(bs)
			if err != nil {
				t.Error(err)
			}
			if err = xl.Render(tt.args.ctx, tt.args.data); err != nil && tt.wantErr {
				return
			} else if err == nil && !tt.wantErr {
			} else {
				t.Error(fmt.Printf("About error: Except %v != %v", tt.wantErr, err == nil))
			}
			result := xl.Result()
			// ioutil.WriteFile(tt.name+".xlsx", result.Bytes(), os.ModePerm)
			if err = checkExcelHelper(result.Bytes(), tt.wantRes); err != nil {
				t.Error(err)
			}
		})
	}
}
