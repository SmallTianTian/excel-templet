package xlsxt

import (
	"reflect"
	"testing"
)

func Test_toStringKeyMap(t *testing.T) {
	type args struct {
		v interface{}
	}
	tests := []struct {
		name    string
		args    args
		want    map[string]interface{}
		wantErr bool
	}{
		{
			name: "test map[string]interface{}",
			args: args{
				v: map[string]interface{}{
					"a": "b",
				},
			},
			want: map[string]interface{}{
				"a": "b",
			},
			wantErr: false,
		},
		{
			name: "test map[string]string",
			args: args{
				v: map[string]string{
					"a": "b",
				},
			},
			want: map[string]interface{}{
				"a": "b",
			},
			wantErr: false,
		},
		{
			name: "test map[string]int",
			args: args{
				v: map[string]int{
					"a": 1,
				},
			},
			want: map[string]interface{}{
				"a": 1,
			},
			wantErr: false,
		},
		{
			name: "test map[string]map[string]string",
			args: args{
				v: map[string]map[string]string{
					"a": {"b": "c"},
				},
			},
			want: map[string]interface{}{
				"a": map[string]string{
					"b": "c",
				},
			},
			wantErr: false,
		},
		{
			name: "test string not map[string]interface{}",
			args: args{
				v: "this is a string",
			},
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, err := toStringKeyMap(tt.args.v)
			if (err != nil) != tt.wantErr {
				t.Errorf("toStringKeyMap() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(got, tt.want) {
				t.Errorf("toStringKeyMap() = %v, want %v", got, tt.want)
			}
		})
	}
}
