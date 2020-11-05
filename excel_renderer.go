package xlsxt

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"reflect"
	"regexp"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/SmallTianTian/go-tools/slice"
)

const (
	ctxCacheKey    = "_xlsxt_ctx"
	renderCacheKey = "_xlsxt_render_cache"
	sheetDataKey   = "_xlsxt_sheet_data"
)

var (
	rangeRgx    = regexp.MustCompile(`{{range (\w*)}}`)
	rowRangeRgx = regexp.MustCompile(`{{rowRange (\w*)}}`)
)

// 错误码从 20000 开始
var (
	NotStringKeyMapValue = errors.New("code: 20000, Not a string key map value.")
	NotMatchRangeEnd     = errors.New("code: 20001, Range not match end.")
	RenderCancel         = errors.New("code: 20002, range is cancel.")
)

type Xlsxt struct {
	file *excelize.File
	buf  bytes.Buffer

	// private
	ctx          context.Context
	cacheRender  map[string]*Parse
	sheetData    map[string]interface{}
	curSheetData map[string]interface{}
}

func NewFromBinary(content []byte) (res *Xlsxt, err error) {
	f, err := excelize.OpenReader(bytes.NewReader(content))
	if err != nil {
		return nil, err
	}
	return &Xlsxt{file: f, cacheRender: make(map[string]*Parse)}, nil
}

// Render renders report and stores it in a struct
func (m *Xlsxt) Render(ctx context.Context, in interface{}) (err error) {
	if ctx == nil {
		ctx = context.Background()
	}
	m.ctx = ctx
	skm, err := toStringKeyMap(in)
	if err != nil {
		return err
	}

	m.buf, err = m.defaultRender(skm)
	return
}

func (m *Xlsxt) Result() bytes.Buffer {
	return m.buf
}

func (m *Xlsxt) defaultRender(data map[string]interface{}) (buf bytes.Buffer, err error) {
	f := excelize.NewFile()
	sns := m.file.GetSheetList()
	for _, sn := range sns {
		var (
			baseColWidth     excelize.BaseColWidth
			defaultColWidth  excelize.DefaultColWidth
			defaultRowHeight excelize.DefaultRowHeight
			customHeight     excelize.CustomHeight
			zeroHeight       excelize.ZeroHeight
			thickTop         excelize.ThickTop
			thickBottom      excelize.ThickBottom
		)
		if err = m.file.GetSheetFormatPr(sn, &baseColWidth, &defaultColWidth, &defaultRowHeight, &customHeight, &zeroHeight, &thickTop, &thickBottom); err != nil {
			return
		}
		var ssw *excelize.StreamWriter
		if ssw, err = f.NewStreamWriter(sn); err != nil {
			return
		}
		if err = f.SetSheetFormatPr(sn, &baseColWidth, &defaultColWidth, &defaultRowHeight, &customHeight, &zeroHeight, &thickTop, &thickBottom); err != nil {
			return
		}
		// TODO get all rows height?
		var rowsData [][]string
		if rowsData, err = m.file.GetRows(sn); err != nil {
			return
		}

		if m.sheetData, err = getSheetData(data, sn, sns); err != nil {
			return
		}
		// remove current sheet data in other sheet
		delete(data, sn)
		if _, err = m.renderRows(ssw, rowsData, 0); err != nil {
			return
		}
		if err = ssw.Flush(); err != nil {
			return
		}

	}
	b, e := f.WriteToBuffer()
	return *b, e
}

func (m *Xlsxt) renderRows(write *excelize.StreamWriter, rowsData [][]string, rowOffset int) (renderLine int, err error) {
	var axis string
	for w := 0; w < len(rowsData); {
		if m.ctx.Err() != nil {
			return 0, RenderCancel
		}

		if axis, err = excelize.CoordinatesToCellName(1, renderLine+1+rowOffset); err != nil {
			return
		}
		cells := rowsData[w]
		// empty line
		if len(cells) == 0 {
			write.SetRow(axis, nil)
			renderLine++
			w++
			continue
		}

		// range begin
		if ms := rangeRgx.FindStringSubmatch(cells[0]); len(ms) == 2 {
			rangeKey := ms[1]
			end := getEndRowIndex(rowsData[w+1:])
			// can't find end
			if end == -1 {
				return 0, NotMatchRangeEnd
			}
			// skip range line
			// no valid render line
			if end == 1 {
				w += 2
				continue
			}
			var rl int
			if rl, err = m.renderRangeRow(write, rangeKey, rowsData[w+1:w+end], w+rowOffset); err != nil {
				return
			}
			renderLine += rl
			w += end
			w++
			continue
		}

		// no row range
		rowResultData := make([]interface{}, 0, len(rowsData[w]))
		for _, item := range rowsData[w] {
			var cellResult *excelize.Cell
			if cellResult, err = m.renderCells(item); err != nil {
				return
			}
			rowResultData = append(rowResultData, cellResult)
		}
		write.SetRow(axis, rowResultData)
		renderLine++
		w++
	}
	return
}

func (m *Xlsxt) renderRangeRow(write *excelize.StreamWriter, rangeKey string, rowsData [][]string, offset int) (renderLine int, err error) {
	rangeD, has := m.sheetData[rangeKey]
	// no valid render data
	if !has {
		return len(rowsData), nil
	}
	m.curSheetData = excludeKeyMap(m.sheetData, rangeKey)

	dc := getChanKeyMap(rangeD)
	for i := 0; ; i++ {
		if v, ok := <-dc; !ok {
			break
		} else {
			m.curSheetData = mergeMap(m.curSheetData, v)

			l, err := m.renderRows(write, rowsData, offset)
			if err != nil {
				return 0, err
			}
			renderLine += l
			offset += l
		}
	}
	return
}

func (m *Xlsxt) renderCells(tlp string) (a *excelize.Cell, err error) {

	defer func() {
		if e := recover(); e != nil {
			err = fmt.Errorf("code: 20002, UNKNOW ERR. %v", e)
		}
	}()
	if _, in := m.cacheRender[tlp]; !in {
		if m.cacheRender[tlp], err = NewParse(tlp); err != nil {
			return
		}
	}
	tp := m.cacheRender[tlp]
	var v interface{}

	if v, err = tp.Exec(m.ctx, m.curSheetData); err != nil {
		return
	}
	return &excelize.Cell{Value: v}, nil
}

func getEndRowIndex(rowsData [][]string) int {
	var inStack int
	for index, v := range rowsData {
		if len(v) == 0 {
			continue
		}
		fV := v[0]
		if fV == "{{end}}" {
			if inStack == 0 {
				// {{range }}
				return index + 1
			}
			inStack--
			continue
		}
		if rangeRgx.MatchString(fV) {
			inStack++
		}
	}
	return -1
}

func getSheetData(in map[string]interface{}, sn string, allSN []string) (result map[string]interface{}, err error) {
	if result, err = toStringKeyMap(in[sn]); err != nil {
		return
	}

	noOtherSheetName := excludeKeyMap(in, allSN...)
	return mergeMap(result, noOtherSheetName), nil
}

func excludeKeyMap(m map[string]interface{}, key ...string) map[string]interface{} {
	result := make(map[string]interface{})
	for k, v := range m {
		if !slice.StringInSlice(key, k) {
			result[k] = v
		}
	}
	return result
}

func mergeMap(dist, fresh map[string]interface{}) map[string]interface{} {
	result := make(map[string]interface{})
	for k, v := range dist {
		result[k] = v
	}
	for k, v := range fresh {
		result[k] = v
	}
	return result
}

func getChanKeyMap(v interface{}) <-chan map[string]interface{} {
	if ckm, ok := v.(chan map[string]interface{}); ok {
		return ckm
	}
	c := make(chan map[string]interface{})
	if akm, ok := v.([]map[string]interface{}); ok {
		go func() {
			for _, v := range akm {
				c <- v
			}
			close(c)
		}()
		return c
	}

	var (
		rt reflect.Type
		rv reflect.Value
	)
NoPtr:
	rt = reflect.TypeOf(v)
	rv = reflect.ValueOf(v)
	if rt.Kind() == reflect.Ptr && rv.CanAddr() {
		v = rv.Addr().Interface()
		goto NoPtr
	}

	if rt.Kind() != reflect.Array && rt.Kind() != reflect.Chan && rt.Kind() != reflect.Slice {
		close(c)
		return c
	}

	if rt.Kind() == reflect.Chan {
		go func() {
			for {
				if v, ok := rv.Recv(); ok {
					if skm, err := toStringKeyMap(v.Interface()); err != nil {
						break
					} else {
						c <- skm
					}
				} else {
					break
				}
			}
			close(c)
		}()
		return c
	}

	go func() {
		l := rv.Len()
		for i := 0; i < l; i++ {
			if skm, err := toStringKeyMap(rv.Index(i).Interface()); err != nil {
				break
			} else {
				c <- skm
			}
		}
		close(c)
	}()
	return c
}
