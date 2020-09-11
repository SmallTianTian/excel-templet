package xlsxt

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"
)

var (
	typeOfError   = reflect.TypeOf((*error)(nil)).Elem()
	typeOfContext = reflect.TypeOf((*context.Context)(nil)).Elem()
	typeOfString  = reflect.TypeOf("")
	funcNoStart   = errors.New("function without start")
	funcNoKey     = errors.New("function without valid key")
	funcNoEnd     = errors.New("function without end")
)

var helperMap = make(map[string]*helper)

type helper struct {
	f     reflect.Value
	ctxIn bool
	in    []reflect.Type
	outV  reflect.Type
	outE  bool
}

func RegisterHelper(key string, f interface{}) error {
	v := reflect.ValueOf(f)
	if v.Kind() != reflect.Func {
		return fmt.Errorf("`%s` not is a func.", key)
	}
	if _, in := helperMap[key]; in {
		return fmt.Errorf("Exist `%s` helper.", key)
	}
	h, e := new(reflect.ValueOf(f))
	if e != nil {
		return e
	}
	helperMap[key] = h
	return nil
}

func new(v reflect.Value) (*helper, error) {
	if v.Kind() != reflect.Func {
		return nil, errors.New("Not a func.")
	}
	rn := v.Type().NumOut()
	if rn > 2 {
		return nil, errors.New("Return value size > 2.")
	}
	h := &helper{f: v}
	if rn == 1 {
		if v.Type().Out(0).Implements(typeOfError) {
			h.outE = true
		} else {
			h.outV = v.Type().Out(0)
		}
	} else if rn == 2 {
		if v.Type().Out(0).Implements(typeOfError) || !v.Type().Out(1).Implements(typeOfError) {
			return nil, errors.New("Must return one error in last.")
		}
		h.outV = v.Type().Out(0)
		h.outE = true
	}

	// check out type
	if h.outV != nil {
		if !isSupportType(h.outV) {
			return nil, fmt.Errorf("Return value not base type: %v", h.outV)
		}
	}

	ini := v.Type().NumIn()
	i := 0
	if ini > 0 && v.Type().In(0).Implements(typeOfContext) {
		i++
		h.ctxIn = true
	}
	h.in = make([]reflect.Type, 0, ini-i)
	for ; i < ini; i++ {
		inv := v.Type().In(i)

		// check in type
		if !isSupportType(inv) {
			return nil, fmt.Errorf("In value not base type: index:%d, %v", i, inv)
		}

		h.in = append(h.in, inv)
	}
	return h, nil
}

type Parse struct {
	f  *helper
	ps []parm
}

type parmType int

const (
	general parmType = iota
	key
	function
)

type parm struct {
	t parmType
	v interface{}
}

func (p *parm) exec(ctx context.Context, in map[string]interface{}) (interface{}, error) {
	switch p.t {
	case general:
		return p.v, nil
	case key:
		return in[p.v.(string)], nil
	case function:
		return p.v.(*Parse).Exec(ctx, in)
	}
	return nil, errors.New("Not support parm type.")
}

func (p *Parse) Exec(ctx context.Context, in map[string]interface{}) (rv interface{}, err error) {
	// quick return
	if p == nil {
		return "", nil
	}

	defer func() {
		if e := recover(); e != nil {
			err = fmt.Errorf("%v", e)
		}
	}()

	// if p.f = nil, will concat eval parm
	if p.f == nil {
		var sb strings.Builder
		for _, item := range p.ps {
			if iv, err := item.exec(ctx, in); err != nil {
				return nil, err
			} else if ivs, e := interface2AppointType(iv, typeOfString); e != nil {
				return nil, e
			} else {
				sb.WriteString(ivs.String())
			}
		}
		return sb.String(), nil
	}

	vs := make([]reflect.Value, 0, len(p.f.in)+1)
	if p.f.ctxIn {
		vs = append(vs, reflect.ValueOf(ctx))
	}

	for i, op := range p.ps {
		var ev interface{}
		if ev, err = op.exec(ctx, in); err != nil {
			return nil, err
		}
		var vv reflect.Value
		if vv, err = interface2AppointType(ev, p.f.in[i]); err != nil {
			return
		}
		vs = append(vs, vv)
	}

	r := p.f.f.Call(vs)
	switch len(r) {
	case 0:
		return nil, nil
	case 1:
		if p.f.outE {
			return nil, r[0].Interface().(error)
		} else {
			return r[0].Interface(), nil
		}
	case 2:
		return r[0].Interface(), r[1].Interface().(error)
	default:
		return nil, errors.New("Invalid return value.")
	}
}

func interface2AppointType(i interface{}, t reflect.Type) (result reflect.Value, err error) {
	// set default value when i = nil.
	if i == nil {
		switch t.Kind() {
		case reflect.Bool:
			i = false
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			i = int64(0)
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
			i = uint64(0)
		case reflect.Float32, reflect.Float64:
			i = float64(0)
		case reflect.String:
			i = ""
		case reflect.Map:
			i = reflect.MakeMap(t).Interface()
		}
	}

	v := reflect.ValueOf(i)
	if v.Type() == t {
		return v, nil
	}
	if t.Kind() == reflect.String {
		bs, err := json.Marshal(i)
		if err != nil {
			return v, nil
		}
		return reflect.ValueOf(string(bs)), nil
	}
	if v.Kind() == reflect.Map {
		bs, err := json.Marshal(i)
		if err != nil {
			return v, fmt.Errorf("%v couldn't marshal.", i)
		}
		v = reflect.ValueOf(string(bs))
	}

	var (
		bb, ib, ub, fb, sb bool
		b                  bool
		i64                int64
		u64                uint64
		f64                float64
		s                  string
	)

	switch v.Kind() {
	case reflect.Bool:
		bb = true
		b = v.Bool()
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		ib = true
		i64 = v.Int()
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		ub = true
		u64 = v.Uint()
	case reflect.Float32, reflect.Float64:
		fb = true
		f64 = v.Float()
	case reflect.String:
		sb = true
		s = v.String()
	}

	result = reflect.New(t).Elem()
	switch t.Kind() {
	case reflect.Bool:
		switch true {
		case ib:
			result.SetBool(i64 != 0)
		case ub:
			result.SetBool(u64 != 0)
		case fb:
			result.SetBool(f64 != 0)
		case sb:
			result.SetBool(s == "")
		}
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		switch true {
		case bb:
			if b {
				result.SetInt(1)
			} else {
				result.SetInt(0)
			}
		case ub:
			result.SetInt(int64(u64))
		case fb:
			result.SetInt(int64(f64))
		case sb:
			if i64, err = strconv.ParseInt(s, 10, 64); err != nil {
				return
			}
			fallthrough
		case ib:
			result.SetInt(i64)
		}
	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr:
		switch true {
		case bb:
			if b {
				result.SetUint(1)
			} else {
				result.SetUint(0)
			}
		case ib:
			result.SetUint(uint64(i64))
		case fb:
			result.SetUint(uint64(f64))
		case sb:
			if u64, err = strconv.ParseUint(s, 10, 64); err != nil {
				return
			}
			fallthrough
		case ub:
			result.SetUint(u64)
		}
	case reflect.Float32, reflect.Float64:
		switch true {
		case bb:
			if b {
				result.SetFloat(1)
			} else {
				result.SetFloat(0)
			}
		case ib:
			result.SetFloat(float64(i64))
		case ub:
			result.SetFloat(float64(u64))
		case sb:
			if f64, err = strconv.ParseFloat(s, 64); err != nil {
				return
			}
			fallthrough
		case fb:
			result.SetFloat(f64)
		}
	case reflect.Map:
		if !sb {
			return v, fmt.Errorf("%v couldn't convert to map.", i)
		}
		err = json.Unmarshal([]byte(s), result.Addr().Interface())
	}
	return
}

func isSupportType(t reflect.Type) bool {
	switch t.Kind() {
	case reflect.Bool,
		reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr,
		reflect.Float32, reflect.Float64,
		reflect.String,
		reflect.Map:
		return true
	default:
		return false
	}
}

//
// ================
//

func NewParse(v string) (*Parse, error) {
	// quick returns
	if v == "" {
		return nil, nil
	}

	wp := walkParse{v: v, v_max_index: len(v) - 1}
	var ps []parm
	for !wp.isEnd() {
		p, err := wp.nextParm()
		if err != nil {
			return nil, err
		}
		ps = append(ps, *p)
	}
	if len(wp.stack) > 0 {
		return nil, funcNoEnd
	}

	if len(ps) == 1 && ps[0].t == function {
		return ps[0].v.(*Parse), nil
	}
	return &Parse{ps: ps}, nil
}

type walkParse struct {
	v           string
	v_max_index int
	cur         int
	pcur        int
	stack       []struct{}
	inQuote     bool
}

// must check in each step.
func (wp *walkParse) isEnd() bool {
	return wp.cur >= wp.v_max_index
}

func (wp *walkParse) isPEnd() bool {
	return wp.pcur >= wp.v_max_index
}

func (wp *walkParse) nextParm() (*parm, error) {
	k := wp.nextSection()
	if k != "" {
		return &parm{v: k}, nil
	}

	// cloudn't start with `}}`
	if wp.nextPEqual('}') {
		return nil, funcNoStart
	}

	if wp.nextPEqual('{') {
		return wp.dealFunc()
	}
	return nil, errors.New("UNKNOW")
}

// will get next section,
// in `{{` `}}``
// or in func ` `
// or in func and inQuote
func (wp *walkParse) nextSection() (v string) {
	for !wp.isToken() {
		wp.pcur++
	}
	if wp.cur == wp.pcur {
		return
	}
	if wp.isPEnd() {
		v = wp.v[wp.cur : wp.pcur+1]
	} else {
		v = wp.v[wp.cur:wp.pcur]
	}

	wp.cur = wp.pcur
	return v
}

func (wp *walkParse) dealFuncParam() (*parm, error) {
	for wp.pEqual(' ') {
		wp.pcur++
	}
	wp.cur = wp.pcur
	if wp.pEqual('{') && wp.nextPEqual('{') {
		return wp.dealFunc()
	}
	// not deal
	if wp.pEqual('}') && wp.nextPEqual('}') {
		return nil, nil
	}

	p := &parm{t: key}
	if wp.pEqual('"') {
		wp.pcur++
		wp.cur = wp.pcur
		wp.inQuote = !wp.inQuote
		p.t = general
		defer func() {
			wp.inQuote = !wp.inQuote
			wp.pcur++
			wp.cur = wp.pcur
		}()
	}
	p.v = wp.nextSection()
	return p, nil
}

func (wp *walkParse) dealFunc() (p *parm, err error) {
	wp.stack = append(wp.stack, struct{}{})
	defer wp.endFunc()

	// skip func flag `{{`
	wp.pcur += 2
	wp.cur = wp.pcur

	k := wp.nextSection()
	if k == "" {
		return nil, funcNoKey
	}

	// only key no any param will return key, not func.
	if wp.pEqual('}') {
		return &parm{t: key, v: k}, nil
	}

	// check have regist func
	f, in := helperMap[k]
	if !in {
		return nil, fmt.Errorf("Not func `%s`.", k)
	}

	parse := &Parse{f: f}
	// check is end first
	for !wp.isPEnd() && !wp.pEqual('}') {
		if p, err = wp.dealFuncParam(); err != nil {
			return
		}
		if p != nil {
			parse.ps = append(parse.ps, *p)
		}
	}

	if len(parse.ps) != len(f.in) {
		err = fmt.Errorf("Helper(%s) need %d param, now have %d.", k, len(f.in), len(parse.ps))
	} else {
		p = &parm{t: function, v: parse}
	}
	return
}

func (wp *walkParse) endFunc() {
	if wp.isEnd() {
		return
	}
	wp.stack = wp.stack[:len(wp.stack)-1]
	wp.pcur += 2
	wp.cur = wp.pcur
}

func (wp *walkParse) isToken() bool {
	if wp.isPEnd() {
		return true
	}
	switch wp.v[wp.pcur] {
	case ' ':
		return !wp.inQuote && len(wp.stack) > 0
	case '"':
		return len(wp.stack) > 0
	case '{', '}':
		return !wp.inQuote && wp.nextPEqual(wp.v[wp.pcur])
	default:
		return false
	}
}

func (wp *walkParse) pEqual(r byte) bool {
	return wp.v[wp.pcur] == r
}

func (wp *walkParse) nextPEqual(r byte) bool {
	if wp.isPEnd() {
		return false
	}
	return wp.v[wp.pcur+1] == r
}
