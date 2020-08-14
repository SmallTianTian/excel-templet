package xlsxt

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"reflect"
	"regexp"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/SmallTianTian/go-tools/slice"
	"github.com/aymerick/raymond"
)

const (
	randerCacheKey = "_xlsxt_rander_cache"
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
}

func NewFromBinary(content []byte) (res *Xlsxt, err error) {
	f, err := excelize.OpenReader(bytes.NewReader(content))
	if err != nil {
		return nil, err
	}
	return &Xlsxt{file: f}, nil
}

// Render renders report and stores it in a struct
func (m *Xlsxt) Render(ctx context.Context, in interface{}) (err error) {
	if ctx == nil {
		ctx = context.Background()
	}
	skm, err := toStringKeyMap(in)
	if err != nil {
		return err
	}

	m.buf, err = defaultRender(ctx, m.file, skm)
	return
}

func (m *Xlsxt) Result() bytes.Buffer {
	return m.buf
}

func defaultRender(ctx context.Context, temp *excelize.File, data map[string]interface{}) (buf bytes.Buffer, err error) {
	f := excelize.NewFile()
	sns := temp.GetSheetList()
	ctx = context.WithValue(ctx, randerCacheKey, make(map[string]*raymond.Template))
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
		if err = temp.GetSheetFormatPr(sn, &baseColWidth, &defaultColWidth, &defaultRowHeight, &customHeight, &zeroHeight, &thickTop, &thickBottom); err != nil {
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
		if rowsData, err = temp.GetRows(sn); err != nil {
			return
		}

		var sheetData map[string]interface{}
		if sheetData, err = getSheetData(data, sn, sns); err != nil {
			return
		}
		sheetCtx := context.WithValue(ctx, sheetDataKey, sheetData)
		// remove current sheet data in other sheet
		delete(data, sn)
		if _, err = renderRows(sheetCtx, ssw, rowsData, 0); err != nil {
			return
		}
		if err = ssw.Flush(); err != nil {
			return
		}

	}
	b, e := f.WriteToBuffer()
	return *b, e
}

func renderRows(ctx context.Context, write *excelize.StreamWriter, rowsData [][]string, rowOffset int) (renderLine int, err error) {
	var axis string
	for w := 0; w < len(rowsData); {
		if ctx.Err() != nil {
			return 0, RenderCancel
		}

		if axis, err = excelize.CoordinatesToCellName(1, w+1+rowOffset); err != nil {
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
			if rl, err := renderRangeRow(ctx, write, rangeKey, rowsData[w+1:w+end], w+rowOffset); err != nil {
				return 0, nil
			} else {
				renderLine += rl
				rowOffset += rl
			}
			w += end
			w++
			continue
		}

		// no row range
		rowResultData := make([]interface{}, 0, len(rowsData[w]))
		for _, item := range rowsData[w] {
			var cellResult *excelize.Cell
			if cellResult, err = renderCells(ctx, item); err != nil {
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

func renderRangeRow(ctx context.Context, write *excelize.StreamWriter, rangeKey string, rowsData [][]string, offset int) (renderLine int, err error) {
	sD := ctx.Value(sheetDataKey).(map[string]interface{})
	rangeD, has := sD[rangeKey]
	// no valid render data
	if !has {
		return len(rowsData), nil
	}
	ori := excludeKeyMap(sD, rangeKey)

	dc := getChanKeyMap(rangeD)
	for i := 0; ; i++ {
		if v, ok := <-dc; !ok {
			break
		} else {
			ctx = context.WithValue(ctx, sheetDataKey, mergeMap(ori, v))

			l, err := renderRows(ctx, write, rowsData, offset)
			if err != nil {
				return 0, nil
			}
			renderLine += l
			offset += l
		}
	}
	return
}

func renderCells(ctx context.Context, tlp string) (a *excelize.Cell, err error) {
	cacheRender := ctx.Value(randerCacheKey).(map[string]*raymond.Template)
	sD := ctx.Value(sheetDataKey).(map[string]interface{})
	defer func() {
		if e := recover(); e != nil {
			err = fmt.Errorf("code: 20002, UNKNOW ERR. %v", e)
		}
	}()
	if _, in := cacheRender[tlp]; !in {
		cacheRender[tlp] = raymond.MustParse(tlp)
	}
	tp := cacheRender[tlp].Clone()
	var v string
	if v, err = tp.Exec(sD); err != nil {
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
