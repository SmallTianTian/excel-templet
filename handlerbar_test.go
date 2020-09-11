package xlsxt

import (
	"context"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"testing"
)

func cleanHelpers() {
	helperMap = make(map[string]*helper)
}

func wrapHelper(i interface{}) *helper {
	h, _ := new(reflect.ValueOf(i))
	return h
}

func equalParse(have, want *Parse) error {
	if !(have == nil && want == nil) && (have == nil || want == nil) {
		return fmt.Errorf("Not both nil, have: %v, want: %v", have, want)
	}
	if !(have.f == nil && want.f == nil) && (have.f == nil || want.f == nil) {
		return fmt.Errorf("Func not both nil, have: %v, want: %v", have.f, want.f)
	}
	if len(have.ps) != len(want.ps) {
		return fmt.Errorf("Param ps len not equal, have: %v, want: %v", len(have.ps), len(want.ps))
	}
	for i, h := range have.ps {
		if h.t != want.ps[i].t {
			return fmt.Errorf("Param ps type not equal, in: %d, have: %v, want: %v", i, h.t, want.ps[i].t)
		}
		if h.t != function {
			if h.v != want.ps[i].v {
				return fmt.Errorf("Param ps value not equal, in: %d, have: %v, want: %v", i, h.v, want.ps[i].v)
			}
		} else {
			if err := equalParse(h.v.(*Parse), want.ps[i].v.(*Parse)); err != nil {
				return fmt.Errorf("Param ps value not equal, in: %d, have: %v, want: %v, err: %v", i, h.v, want.ps[i].v, err)
			}
		}
	}
	return nil
}

func equalHelper(have, want *helper) error {
	if have.ctxIn != want.ctxIn || have.outE != want.outE || ((have.outV == nil || want.outV == nil) && want.outV != have.outV) ||
		(have.outV != nil && have.outV.Kind() != want.outV.Kind()) {
		return fmt.Errorf("(ctxIn, outE, outV): (%v, %v, %v), want (%v, %v, %v).",
			have.ctxIn, have.outE, have.outV.Kind(), want.ctxIn, want.outE, want.outV.Kind())
	}
	if have.f.Kind() != want.f.Kind() {
		return fmt.Errorf("f = %v, wantF %v.", have.f.Kind(), want.f.Kind())
	}
	if len(have.in) != len(want.in) {
		return fmt.Errorf("len in = %v, want %v.", len(have.in), len(want.in))
	}
	for i, item := range have.in {
		if item.Kind() != want.in[i].Kind() {
			return fmt.Errorf("%d's in val = %v, want %v.", i, item.Kind(), want.in[i].Kind())
		}
	}
	return nil
}

func TestRegisterHelper(t *testing.T) {
	type args struct {
		key string
		f   interface{}
	}
	tests := []struct {
		name    string
		args    args
		wantErr bool
		wantVal *helper
	}{
		{
			name: "no in no return",
			args: args{
				key: "Test",
				f:   func() {},
			},
			wantVal: &helper{
				f: reflect.ValueOf(func() {}),
			},
		},
		{
			name: "no in value return",
			args: args{
				key: "Test",
				f:   func() string { return "a" },
			},
			wantVal: &helper{
				f:    reflect.ValueOf(func() string { return "a" }),
				outV: reflect.TypeOf(""),
			},
		},
		{
			name: "no in error return",
			args: args{
				key: "Test",
				f:   func() error { return nil },
			},
			wantVal: &helper{
				f:    reflect.ValueOf(func() error { return nil }),
				outE: true,
			},
		},
		{
			name: "no in value and error return",
			args: args{
				key: "Test",
				f:   func() (int, error) { return 0, nil },
			},
			wantVal: &helper{
				f:    reflect.ValueOf(func() (int, error) { return 0, nil }),
				outE: true,
				outV: reflect.TypeOf(0),
			},
		},
		{
			name: "only ctx in no return",
			args: args{
				key: "Test",
				f:   func(context.Context) {},
			},
			wantVal: &helper{
				f:     reflect.ValueOf(func(context.Context) {}),
				ctxIn: true,
			},
		},
		{
			name: "one in no return",
			args: args{
				key: "Test",
				f:   func(int) {},
			},
			wantVal: &helper{
				f:  reflect.ValueOf(func(int) {}),
				in: []reflect.Type{reflect.TypeOf(0)},
			},
		},
		{
			name: "ctx and one in no return",
			args: args{
				key: "Test",
				f:   func(context.Context, int) {},
			},
			wantVal: &helper{
				f:     reflect.ValueOf(func(context.Context, int) {}),
				ctxIn: true,
				in:    []reflect.Type{reflect.TypeOf(0)},
			},
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			cleanHelpers()
			if err := RegisterHelper(tt.args.key, tt.args.f); (err != nil) != tt.wantErr {
				t.Errorf("RegisterHelper() error = %v, wantErr %v", err, tt.wantErr)
			}
			v := helperMap[tt.args.key]
			if v == nil {
				t.Errorf("RegisterHelper() should have helper about key %s, now is nil", tt.args.key)
			}
			if tt.wantVal == nil {
				t.Errorf("RegisterHelper() value = %v, wantVal nil.", v)
			}
			if err := equalHelper(v, tt.wantVal); err != nil {
				t.Errorf("RegisterHelper() helper = %v, want %v, Err: %v", v, tt.wantVal, err)
			}
		})
	}
}

func Test_interface2AppointType(t *testing.T) {
	type args struct {
		i interface{}
		t reflect.Type
	}
	tests := []struct {
		name       string
		args       args
		wantResult reflect.Value
		wantErr    bool
	}{
		{
			name: "bool2bool",
			args: args{
				i: false,
				t: reflect.TypeOf(true),
			},
			wantResult: reflect.ValueOf(false),
		},
		{
			name: "bool2int8-false",
			args: args{
				i: false,
				t: reflect.TypeOf(int8(0)),
			},
			wantResult: reflect.ValueOf(int8(0)),
		},
		{
			name: "bool2int8-true",
			args: args{
				i: true,
				t: reflect.TypeOf(int8(0)),
			},
			wantResult: reflect.ValueOf(int8(1)),
		},
		{
			name: "bool2int16-false",
			args: args{
				i: false,
				t: reflect.TypeOf(int16(0)),
			},
			wantResult: reflect.ValueOf(int16(0)),
		},
		{
			name: "bool2int16-true",
			args: args{
				i: true,
				t: reflect.TypeOf(int16(0)),
			},
			wantResult: reflect.ValueOf(int16(1)),
		},
		{
			name: "bool2int32-false",
			args: args{
				i: false,
				t: reflect.TypeOf(int32(0)),
			},
			wantResult: reflect.ValueOf(int32(0)),
		},
		{
			name: "bool2int32-true",
			args: args{
				i: true,
				t: reflect.TypeOf(int32(0)),
			},
			wantResult: reflect.ValueOf(int32(1)),
		},
		{
			name: "bool2int64-false",
			args: args{
				i: false,
				t: reflect.TypeOf(int64(0)),
			},
			wantResult: reflect.ValueOf(int64(0)),
		},
		{
			name: "bool2int64-true",
			args: args{
				i: true,
				t: reflect.TypeOf(int64(0)),
			},
			wantResult: reflect.ValueOf(int64(1)),
		},
		{
			name: "bool2int-false",
			args: args{
				i: false,
				t: reflect.TypeOf(int(0)),
			},
			wantResult: reflect.ValueOf(int(0)),
		},
		{
			name: "bool2int-true",
			args: args{
				i: true,
				t: reflect.TypeOf(int(0)),
			},
			wantResult: reflect.ValueOf(int(1)),
		},
		{
			name: "bool2uint8-false",
			args: args{
				i: false,
				t: reflect.TypeOf(uint8(0)),
			},
			wantResult: reflect.ValueOf(uint8(0)),
		},
		{
			name: "bool2uint8-true",
			args: args{
				i: true,
				t: reflect.TypeOf(uint8(0)),
			},
			wantResult: reflect.ValueOf(uint8(1)),
		},
		{
			name: "bool2uint16-false",
			args: args{
				i: false,
				t: reflect.TypeOf(uint16(0)),
			},
			wantResult: reflect.ValueOf(uint16(0)),
		},
		{
			name: "bool2uint16-true",
			args: args{
				i: true,
				t: reflect.TypeOf(uint16(0)),
			},
			wantResult: reflect.ValueOf(uint16(1)),
		},
		{
			name: "bool2uint32-false",
			args: args{
				i: false,
				t: reflect.TypeOf(uint32(0)),
			},
			wantResult: reflect.ValueOf(uint32(0)),
		},
		{
			name: "bool2uint32-true",
			args: args{
				i: true,
				t: reflect.TypeOf(uint32(0)),
			},
			wantResult: reflect.ValueOf(uint32(1)),
		},
		{
			name: "bool2uint64-false",
			args: args{
				i: false,
				t: reflect.TypeOf(uint64(0)),
			},
			wantResult: reflect.ValueOf(uint64(0)),
		},
		{
			name: "bool2uint64-true",
			args: args{
				i: true,
				t: reflect.TypeOf(uint64(0)),
			},
			wantResult: reflect.ValueOf(uint64(1)),
		},
		{
			name: "bool2uint-false",
			args: args{
				i: false,
				t: reflect.TypeOf(uint(0)),
			},
			wantResult: reflect.ValueOf(uint(0)),
		},
		{
			name: "bool2uint-true",
			args: args{
				i: true,
				t: reflect.TypeOf(uint(0)),
			},
			wantResult: reflect.ValueOf(uint(1)),
		},
		{
			name: "bool2float32-false",
			args: args{
				i: false,
				t: reflect.TypeOf(float32(0)),
			},
			wantResult: reflect.ValueOf(float32(0)),
		},
		{
			name: "bool2float32-true",
			args: args{
				i: true,
				t: reflect.TypeOf(float32(0)),
			},
			wantResult: reflect.ValueOf(float32(1)),
		},
		{
			name: "bool2float64-false",
			args: args{
				i: false,
				t: reflect.TypeOf(float64(0)),
			},
			wantResult: reflect.ValueOf(float64(0)),
		},
		{
			name: "bool2float64-true",
			args: args{
				i: true,
				t: reflect.TypeOf(float64(0)),
			},
			wantResult: reflect.ValueOf(float64(1)),
		},
		{
			name: "bool2string-false",
			args: args{
				i: false,
				t: reflect.TypeOf(""),
			},
			wantResult: reflect.ValueOf("false"),
		},
		{
			name: "bool2string-true",
			args: args{
				i: true,
				t: reflect.TypeOf(""),
			},
			wantResult: reflect.ValueOf("true"),
		},
		{
			name: "nil2map[string]int",
			args: args{
				t: reflect.TypeOf(map[string]int{}),
			},
			wantResult: reflect.ValueOf(map[string]int{}),
		},
		{
			name: "nil2map[int]string",
			args: args{
				t: reflect.TypeOf(map[int]string{}),
			},
			wantResult: reflect.ValueOf(map[int]string{}),
		},
		{
			name: "int82map[string]string",
			args: args{
				i: int8(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "int162map[string]string",
			args: args{
				i: int16(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "int322map[string]string",
			args: args{
				i: int32(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "int642map[string]string",
			args: args{
				i: int64(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "int2map[string]string",
			args: args{
				i: int(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uint82map[string]string",
			args: args{
				i: uint8(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uint162map[string]string",
			args: args{
				i: uint16(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uint322map[string]string",
			args: args{
				i: uint32(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uint642map[string]string",
			args: args{
				i: uint64(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uint2map[string]string",
			args: args{
				i: uint(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "uintptr2map[string]string",
			args: args{
				i: uintptr(8),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "bool2map[string]string",
			args: args{
				i: true,
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "float322map[string]string",
			args: args{
				i: float32(1.23),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "float642map[string]string",
			args: args{
				i: float64(1.23),
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "not_json_string2map[string]string",
			args: args{
				i: "this is not json.",
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "not_match_type_json_string2map[string]string",
			args: args{
				i: `{"key":123}`,
				t: reflect.TypeOf(map[string]string{}),
			},
			wantErr: true,
		},
		{
			name: "json_string2map[string]int",
			args: args{
				i: `{"key":123}`,
				t: reflect.TypeOf(map[string]int{}),
			},
			wantResult: reflect.ValueOf(map[string]int{"key": 123}),
		},
		{
			name: "not_match_type_map2map[string]int",
			args: args{
				i: map[string]string{"key": "123"},
				t: reflect.TypeOf(map[string]int{}),
			},
			wantErr: true,
		},
		{
			name: "map2map[string]string",
			args: args{
				i: map[string]string{"key": "123"},
				t: reflect.TypeOf(map[string]string{}),
			},
			wantResult: reflect.ValueOf(map[string]string{"key": "123"}),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotResult, err := interface2AppointType(tt.args.i, tt.args.t)
			if err != nil && tt.wantErr {
				return
			}
			if (err != nil) != tt.wantErr {
				t.Errorf("interface2AppointType() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(gotResult.Interface(), tt.wantResult.Interface()) {
				t.Errorf("interface2AppointType() = (%T)%v, want (%T)%v", gotResult.Interface(), gotResult.Interface(), tt.wantResult.Interface(), tt.wantResult.Interface())
			}
		})
	}
}

func Test_isSupportType(t *testing.T) {
	type args struct{}
	i := 1
	tests := []struct {
		name string
		args interface{}
		want bool
	}{
		{
			name: "bool",
			args: false,
			want: true,
		},
		{
			name: "int8",
			args: int8(1),
			want: true,
		},
		{
			name: "int16",
			args: int16(1),
			want: true,
		},
		{
			name: "int32",
			args: int32(1),
			want: true,
		},
		{
			name: "int64",
			args: int64(1),
			want: true,
		},
		{
			name: "int",
			args: int(1),
			want: true,
		},
		{
			name: "uint8",
			args: uint8(1),
			want: true,
		},
		{
			name: "uint16",
			args: uint16(1),
			want: true,
		},
		{
			name: "uint32",
			args: uint32(1),
			want: true,
		},
		{
			name: "uint64",
			args: uint64(1),
			want: true,
		},
		{
			name: "uint",
			args: uint(1),
			want: true,
		},
		{
			name: "uintptr",
			args: uintptr(1),
			want: true,
		},
		{
			name: "float32",
			args: float32(1),
			want: true,
		},
		{
			name: "float64",
			args: float64(1),
			want: true,
		},
		{
			name: "string",
			args: "string",
			want: true,
		},
		{
			name: "map",
			args: map[string]int{},
			want: true,
		},
		{
			name: "map int",
			args: map[int]int{},
			want: true,
		},
		{
			name: "slice not!!!",
			args: make([]int, 0),
		},
		{
			name: "array not!!!",
			args: [0]int{},
		},
		{
			name: "struct not!!!",
			args: args{},
		},
		{
			name: "ptr not!!!",
			args: &i,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := isSupportType(reflect.TypeOf(tt.args)); got != tt.want {
				t.Errorf("isSupportType() = %v, want %v", got, tt.want)
			}
		})
	}
}

func TestParse_Exec(t *testing.T) {
	c := context.WithValue(context.Background(), "k", "v")
	type fields struct {
		f  *helper
		ps []parm
	}
	type args struct {
		ctx context.Context
		in  map[string]interface{}
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantRv  interface{}
		wantErr bool
	}{
		{
			name: "no in retun string",
			fields: fields{
				f:  wrapHelper(func() string { return "string" }),
				ps: []parm{},
			},
			wantRv: "string",
		},
		{
			name: "one string in retun string",
			fields: fields{
				f:  wrapHelper(func(a string) string { return "string" + a }),
				ps: []parm{{t: general, v: "s"}},
			},
			wantRv: "strings",
		},
		{
			name: "no in retun error",
			fields: fields{
				f:  wrapHelper(func() error { return errors.New("text string") }),
				ps: []parm{},
			},
			wantErr: true,
		},
		{
			name: "int in retun string",
			fields: fields{
				f:  wrapHelper(func(i int) string { return strconv.Itoa(i) }),
				ps: []parm{{t: general, v: 10000}},
			},
			wantRv: "10000",
		},
		{
			name: "no in no return with panic",
			fields: fields{
				f:  wrapHelper(func() { panic("test") }),
				ps: []parm{},
			},
			wantErr: true,
		},
		{
			name: "ctx in ctx return",
			fields: fields{
				f:  wrapHelper(func(ctx context.Context) bool { return ctx == c }),
				ps: []parm{},
			},
			args: args{
				ctx: c,
			},
			wantRv: true,
		},
		{
			name: "int in int return with string param",
			fields: fields{
				f:  wrapHelper(func(i int) int { return i }),
				ps: []parm{{t: general, v: "10"}},
			},
			wantRv: 10,
		},
		{
			name: "no in int return with return string func",
			fields: fields{
				f: wrapHelper(func(i int) int { return i }),
				ps: []parm{{t: function, v: &Parse{
					f: wrapHelper(func() string { return "10" }),
				}}},
			},
			wantRv: 10,
		},
		{
			name: "no f in parse will concat each param",
			fields: fields{
				ps: []parm{
					{t: function, v: &Parse{
						f: wrapHelper(func() string { return "string" }),
					}},
					{t: general, v: 10},
					{t: general, v: false},
				},
			},
			wantRv: "string10false",
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			p := &Parse{
				f:  tt.fields.f,
				ps: tt.fields.ps,
			}
			gotRv, err := p.Exec(tt.args.ctx, tt.args.in)
			if (err != nil) != tt.wantErr {
				t.Errorf("Parse.Exec() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if !reflect.DeepEqual(gotRv, tt.wantRv) {
				t.Errorf("Parse.Exec() = %v, want %v", gotRv, tt.wantRv)
			}
		})
	}
}

func TestNewParse(t *testing.T) {
	tests := []struct {
		name    string
		args    string
		help    map[string]interface{}
		want    *Parse
		wantErr bool
	}{
		{
			name: "just string",
			args: "string",
			want: &Parse{ps: []parm{{t: general, v: "string"}}},
		},
		{
			name: "string with blank",
			args: "this is string",
			want: &Parse{ps: []parm{{t: general, v: "this is string"}}},
		},
		{
			name: "single brace",
			args: "{single",
			want: &Parse{ps: []parm{{t: general, v: "{single"}}},
		},
		{
			name: "single brace with blank",
			args: "single } string",
			want: &Parse{ps: []parm{{t: general, v: "single } string"}}},
		},
		{
			name: "key",
			args: "{{key}}",
			want: &Parse{ps: []parm{{t: key, v: "key"}}},
		},
		{
			name: "function with single string in",
			args: "{{exist key}}",
			help: map[string]interface{}{"exist": func(s string) {}},
			want: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{t: key, v: "key"}}},
		},
		{
			name: "function with two string in",
			args: "{{exist key value}}",
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: key, v: "value"}}},
		},
		{
			name: "function with two string in and more blank",
			args: "{{exist    key      value}}",
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: key, v: "value"}}},
		},
		{
			name: "function in with word not key",
			args: `{{exist key "value"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: general, v: "value"}}},
		},
		{
			name: "function in with word not key and blank",
			args: `{{exist       key       "value"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: general, v: "value"}}},
		},
		{
			name: "function in with blank word not key",
			args: `{{exist key " value"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: general, v: " value"}}},
		},
		{
			name: "function in with blank&{} word not key",
			args: `{{exist key " value {{}  } a"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "key"}, {t: general, v: " value {{}  } a"}}},
		},
		{
			name: "function first in is blank&{} word not key",
			args: `{{exist " value {{}  } a" key}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: " value {{}  } a"}, {t: key, v: "key"}}},
		},
		{
			name: "func paralleling func",
			args: `{{exist "a" k}}{{exist b "v"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{ps: []parm{
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: "a"}, {t: key, v: "k"}}}},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
			}},
		},
		{
			name: "func paralleling func with blank",
			args: `{{exist "a" k}}   {{exist b "v"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{ps: []parm{
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: "a"}, {t: key, v: "k"}}}},
				{t: general, v: "   "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
			}},
		},
		{
			name: "five func paralleling with blank",
			args: `{{exist "a" k}}   {{exist b "v"}}   {{exist b "v"}}   {{exist b "v"}}   {{exist b "v"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{ps: []parm{
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: "a"}, {t: key, v: "k"}}}},
				{t: general, v: "   "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
				{t: general, v: "   "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
				{t: general, v: "   "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
				{t: general, v: "   "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
			}},
		},
		{
			name: "func paralleling func with key",
			args: `{{exist "a" k}}{{key}}{{exist b "v"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{ps: []parm{
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: "a"}, {t: key, v: "k"}}}},
				{t: key, v: "key"},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
			}},
		},
		{
			name: "func paralleling func with blank and word",
			args: `{{exist "a" k}} key {{exist b "v"}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{ps: []parm{
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: general, v: "a"}, {t: key, v: "k"}}}},
				{t: general, v: " key "},
				{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
			}},
		},
		{
			name: "func nest func",
			args: `{{exist {{exist b "v"}} k}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{
				f: wrapHelper(func(s, s1 string) {}),
				ps: []parm{
					{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "b"}, {t: general, v: "v"}}}},
					{t: key, v: "k"},
				}},
		},
		{
			name: "five func nest",
			args: `{{exist {{exist {{exist {{exist {{exist_2 d}} "c"}} "k"}} "v"}} {{key}}}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}, "exist_2": func(s string) {}},
			want: &Parse{
				f: wrapHelper(func(s, s1 string) {}),
				ps: []parm{
					{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{
						{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{
							{t: function, v: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{
								{t: function, v: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{t: key, v: "d"}}}},
								{t: general, v: "c"}}},
							},
							{t: general, v: "k"}}},
						},
						{t: general, v: "v"}}},
					},
					{t: key, v: "key"},
				}},
		},
		{
			name: "empty expre",
			args: ``,
			want: nil,
		},
		{
			name:    "no func",
			args:    `{{not_exist "a"}}`,
			wantErr: true,
		},

		{
			name:    "func no end",
			args:    `{{exist "a"`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name:    "func no start",
			args:    `exist "a"}}`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name:    "func duplicate start",
			args:    `{{{{exist "a"}}`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name:    "func duplicate end",
			args:    `{{exist "a"}}}}`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name:    "nest func duplicate without end",
			args:    `{{exist {{exist "a"}}`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name:    "nest func duplicate without start",
			args:    `{{exist exist "a"}}}}`,
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
		{
			name: "function first key with single {",
			args: `{{exist {key key}}`,
			help: map[string]interface{}{"exist": func(s, s1 string) {}},
			want: &Parse{f: wrapHelper(func(s, s1 string) {}), ps: []parm{{t: key, v: "{key"}, {t: key, v: "key"}}},
		},
		{
			name: "function param(value) end with blank",
			args: `{{exist "value" }}`,
			help: map[string]interface{}{"exist": func(s string) {}},
			want: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{t: general, v: "value"}}},
		},
		{
			name: "function param(key) end with blank",
			args: `{{exist key }}`,
			help: map[string]interface{}{"exist": func(s string) {}},
			want: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{t: key, v: "key"}}},
		},
		{
			name: "function param(function) end with blank",
			args: `{{exist {{exist "a"}} }}`,
			help: map[string]interface{}{"exist": func(s string) {}},
			want: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{t: function, v: &Parse{f: wrapHelper(func(s string) {}), ps: []parm{{v: "a"}}}}}},
		},
		{
			name:    "function param size not equal in",
			args:    "{{exist key}}",
			help:    map[string]interface{}{"exist": func(s, s1 string) {}},
			wantErr: true,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if len(tt.help) > 0 {
				cleanHelpers()
				for k, v := range tt.help {
					RegisterHelper(k, v)
				}
			}

			got, err := NewParse(tt.args)
			if (err != nil) != tt.wantErr {
				t.Errorf("NewParse() error = %v, wantErr %v", err, tt.wantErr)
				return
			}
			if got == tt.want && got == nil {
				return
			}
			if err := equalParse(got, tt.want); err != nil {
				t.Errorf("NewParse() have = %v, want %v, err: %v", got.f, tt.want.f, err)
			}
		})
	}
}
