package xlsxt

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"reflect"
	"testing"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
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
	// fmt.Println(have)
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

func Test_RenderBase(t *testing.T) {
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
			name: "render range with paralleling range",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
					{"{{c}}"},
					{"{{range d}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
					{"{{e}}"},
				},
				data: map[string]interface{}{
					"rows": map2InterChanHelper([]interface{}{
						map[string]string{"s": "s1", "d": "d1"},
						map[string]string{"s": "s2", "d": "d2"},
						map[string]string{"s": "s3", "d": "d3"},
					}),
					"c": "d",
					"d": map2InterChanHelper([]interface{}{
						map[string]string{"s": "s4", "d": "d4"},
						map[string]string{"s": "s5", "d": "d5"},
						map[string]string{"s": "s6", "d": "d6"},
					}),
					"e": "f",
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
				{"d"},
				{"string", "s4"},
				{"string", "s5"},
				{"string", "s6"},
				{"f"},
			},
		},
		{
			name: "render range and no range date.",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{s}}`},
					{"{{end}}"},
					{"{{c}}"},
				},
				data: map[string]interface{}{
					"rows": map2InterChanHelper([]interface{}{
						map[string]string{"s": "s1", "d": "d1"},
						map[string]string{"s": "s2", "d": "d2"},
						map[string]string{"s": "s3", "d": "d3"},
					}),
					"c": "d",
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "s1"},
				{"string", "s2"},
				{"string", "s3"},
				{"d"},
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

func Test_RenderFunc(t *testing.T) {
	ctx := context.WithValue(context.Background(), "single", true)
	type args struct {
		ctx  context.Context
		temp [][]string
		data interface{}
	}
	tests := []struct {
		name    string
		helper  map[string]interface{}
		args    args
		wantRes [][]string
		wantErr bool
	}{
		{
			name:   "Render with func and no data.",
			helper: map[string]interface{}{"exist": func(s string) string { return "true" }},
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{exist "c"}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{}, {}, {},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "true"},
				{"string", "true"},
				{"string", "true"},
			},
		},
		{
			name: "Render with func.",
			helper: map[string]interface{}{"is": func(flag, t, f, value string) string {
				if value == flag {
					return t
				}
				return f
			}},
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{is "hello" "t" "f" k}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"k": "hello"}, {"k": "hi"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "t"},
				{"string", "f"},
			},
		},
		{
			name: "Render with return error func",
			helper: map[string]interface{}{"error": func(key string) (string, error) {
				return "", errors.New("error")
			}},
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{error k}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"k": "hello"},
					},
				},
			},
			wantErr: true,
		},
		{
			name: "Render with no specific func",
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{no_func k}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"k": "hello"},
					},
				},
			},
			wantErr: true,
		},
		{
			name:   "Render with func panic",
			helper: map[string]interface{}{"error": func(key string) string { panic("error") }},
			args: args{
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{error k}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"k": "hello"},
					},
				},
			},
			wantErr: true,
		},
		{
			name: "check func with same ctx",
			helper: map[string]interface{}{"ctx": func(c context.Context, k string) string {
				if c == ctx {
					return "true"
				}
				return "false"
			}},
			args: args{
				ctx: ctx,
				temp: [][]string{
					{"Test"},
					{"{{range rows}}"},
					{"string", `{{ctx k}}`},
					{"{{end}}"},
				},
				data: map[string]interface{}{
					"rows": []map[string]interface{}{
						{"k": "hello"},
					},
				},
			},
			wantRes: [][]string{
				{"Test"},
				{"string", "true"},
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cleanHelpers()
			if len(tt.helper) > 0 {
				for k, v := range tt.helper {
					RegisterHelper(k, v)
				}
			}

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
