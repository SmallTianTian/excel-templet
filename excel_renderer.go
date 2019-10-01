package xlsxt

import (
	"bytes"
	"fmt"
	"regexp"
	"strings"

	"github.com/aymerick/raymond"
	"github.com/tealeg/xlsx"
)

var (
	rgx              = regexp.MustCompile(`\{\{\s*(\w+)\.\w+\s*\}\}`)
	rangeRgx         = regexp.MustCompile(`\{\{\s*range\s+(\w+)\s*\}\}`)
	rangeEndRgx      = regexp.MustCompile(`\{\{\s*end\s*\}\}`)
	defaultCellStyle = xlsx.MakeStringStyle(xlsx.DefaultFont(), xlsx.DefaultFill(), xlsx.DefaultAlignment(), xlsx.DefaultBorder())
	nullCell         = &streamCell{
		CellStyle: &defaultCellStyle,
		CellType:  xlsx.CellTypeString.Ptr(),
	}
)

type Xlsxt struct {
	file *xlsx.File
	buf  bytes.Buffer
}

type streamCell struct {
	CellData  string
	CellStyle *xlsx.StreamStyle
	CellType  *xlsx.CellType
}

func (sc *streamCell) ToXlsx() xlsx.StreamCell {
	return xlsx.NewStreamCell(sc.CellData, *sc.CellStyle, *sc.CellType)
}

func NewFromBinary(content []byte) (res *Xlsxt, err error) {
	var file *xlsx.File
	if file, err = xlsx.OpenBinary(content); err == nil {
		res = NewFromXlsx(file)
	}
	return
}

func NewFromXlsx(file *xlsx.File) *Xlsxt {
	return &Xlsxt{file: file}
}

// Render renders report and stores it in a struct
func (m *Xlsxt) Render(in interface{}) error {
	return m.defaultRender(in)
}

func (m *Xlsxt) defaultRender(in interface{}) error {
	streamBuild := xlsx.NewStreamFileBuilder(&m.buf)

	data := make([]map[int]chan []*streamCell, len(m.file.Sheets))
	defer func() {
		for _, item := range data {
			for _, v := range item {
				close(v)
			}
		}
	}()

	for i, item := range m.file.Sheets {
		max, runner, styles := prepareRender(m.file, i, in)
		styles = append(styles, *nullCell.CellStyle)
		data[i] = map[int]chan []*streamCell{max: runner}

		firstLines := fillStreamCell(<-runner, max)
		strs := make([]string, len(firstLines))
		meta := make([]*xlsx.CellMetadata, len(firstLines))
		for i, item := range firstLines {
			strs[i] = item.CellData
			mt := xlsx.MakeCellMetadata(*item.CellType, *item.CellStyle)
			meta[i] = &mt
		}
		if err := streamBuild.AddSheetWithDefaultColumnMetadata(item.Name, strs, meta); err != nil {
			return err
		}
		streamBuild.AddStreamStyleList(styles)
	}

	sb, err := streamBuild.Build()
	if err != nil {
		return err
	}

	for sheetIndex, sheetData := range data {
		if sheetIndex != 0 {
			if err = sb.NextSheet(); err != nil {
				return err
			}
		}
		for max, sheetChan := range sheetData {
			for {
				rowCells := <-sheetChan
				if rowCells == nil {
					break
				}

				fillCells := fillStreamCell(rowCells, max)
				formatCells := make([]xlsx.StreamCell, len(fillCells))
				for i, cell := range fillCells {
					formatCells[i] = cell.ToXlsx()
				}

				if err = sb.WriteS(formatCells); err != nil {
					return err
				}
			}
		}
	}
	return sb.Close()
}

func (m *Xlsxt) Result() bytes.Buffer {
	return m.buf
}

func getCellStyle(cell *xlsx.Cell) *xlsx.StreamStyle {
	numberFmtId := 0
	font := cell.GetStyle().Font
	fill := cell.GetStyle().Fill
	alig := cell.GetStyle().Alignment
	bord := cell.GetStyle().Border
	style := xlsx.MakeStyle(numberFmtId, &font, &fill, &alig, &bord)
	return &style
}

func render(rows []*xlsx.Row, orgion map[string]interface{}, styles [][]*xlsx.StreamStyle, pipline chan []*streamCell) int {
	lineLen := 0
	for i := 0; i < len(rows); i++ {
		rangeProp := getRangeProp(rows[i])
		if rangeProp != "" {
			end := getRangeEndIndex(rows[i+1:], rangeProp)
			rangeData := getRangeCtx(orgion, rangeProp)
			for _, rd := range rangeData {
				rl := 0
				for end-rl > 0 {
					rl += render(rows[i+1:i+1+end], rd, styles[i+1:], pipline)
				}
			}
			i += end + 1
			lineLen += end + 1
		} else {
			cells := make([]*streamCell, len(rows[i].Cells))
			for j, item := range rows[i].Cells {
				tp := item.Type()
				cells[j] = &streamCell{
					CellData:  item.Value,
					CellStyle: styles[i][j],
					CellType:  &tp,
				}
			}

			lineCells := make([]*streamCell, len(cells))
			for i, c := range cells {
				if lc, err := renderCell(c, orgion); err != nil {
					panic(err)
				} else {
					lineCells[i] = lc
				}

			}
			pipline <- lineCells
			lineLen += 1
		}
	}
	return lineLen
}

func prepareRender(file *xlsx.File, index int, in interface{}) (maxCell int, data chan []*streamCell, style []xlsx.StreamStyle) {
	sheet := file.Sheets[index]
	styles := make([][]*xlsx.StreamStyle, len(file.Sheets[index].Rows))

	// get sheet max cells on a line
	// get all cell style
	for i, row := range sheet.Rows {
		if len(row.Cells) > maxCell {
			maxCell = len(row.Cells)
		}

		lineStyle := make([]*xlsx.StreamStyle, len(row.Cells))
		for j, cell := range row.Cells {
			st := getCellStyle(cell)
			lineStyle[j] = st

			style = append(style, *st)
		}
		styles[i] = lineStyle
	}

	originData := getCtx(in, index)
	data = make(chan []*streamCell, 100*1000)
	go func() {
		render(sheet.Rows, originData, styles, data)
		data <- nil
	}()
	return
}

func getRangeEndIndex(rows []*xlsx.Row, prop string) int {
	expect := fmt.Sprintf("{{end %s}}", prop)

	var nesting int
	for idx := 0; idx < len(rows); idx++ {
		cellValue := rows[idx].Cells[0].Value
		if cellValue == expect {
			return idx
		}

		if len(rows[idx].Cells) == 0 {
			continue
		}

		if rangeEndRgx.MatchString(cellValue) {
			if nesting == 0 {
				return idx
			}

			nesting--
			continue
		}

		if rangeRgx.MatchString(cellValue) {
			nesting++
		}
	}

	return -1
}

func renderCell(cell *streamCell, data map[string]interface{}) (result *streamCell, err error) {
	tpl := strings.Replace(cell.CellData, "{{", "{{{", -1)
	tpl = strings.Replace(tpl, "}}", "}}}", -1)
	var out string

	if count := strings.Count(tpl, "{{{"); count > 0 {
		for i := 0; i < count; i++ {
			l := strings.LastIndex(tpl, "{{{")
			s := strings.Index(tpl, "}}}") + len("}}}")
			tp := tpl[l:s]
			var template *raymond.Template
			if template, err = raymond.Parse(tp); err != nil {
				return
			}
			if out, err = template.Exec(data); err != nil {
				return
			}
			tpl = fmt.Sprintf("%s \"%s\"%s", tpl[:l], out, tpl[s:])
		}
	} else {
		var template *raymond.Template
		if template, err = raymond.Parse(tpl); err != nil {
			return
		}
		if out, err = template.Exec(data); err != nil {
			return
		}
	}

	return &streamCell{
		CellData:  out,
		CellStyle: cell.CellStyle,
		CellType:  cell.CellType,
	}, nil
}

func fillStreamCell(origin []*streamCell, max int) (result []*streamCell) {
	switch {
	case len(origin) == max:
		result = origin
	case len(origin) < max:
		result = make([]*streamCell, len(origin), max)
		copy(result, origin)
		for max-len(result) > 0 {
			result = append(result, nullCell)
		}
	case len(origin) > max:
		panic("StreamCell len > max.")
	}
	return
}

func getCtx(in interface{}, i int) map[string]interface{} {
	if ctx, ok := in.(map[string]interface{}); ok {
		return ctx
	}
	if ctxSlice, ok := in.([]interface{}); ok {
		if len(ctxSlice) > i {
			_ctx := ctxSlice[i]
			if ctx, ok := _ctx.(map[string]interface{}); ok {
				return ctx
			}
		}
		return nil
	}
	return nil
}

func getRangeCtx(ctx map[string]interface{}, prop string) []map[string]interface{} {
	val, ok := ctx[prop]
	if !ok {
		return nil
	}

	var propCtx []interface{}
	if propCtx, ok = val.([]interface{}); !ok {
		return nil
	}
	result := make([]map[string]interface{}, 0, len(propCtx))
	for _, inter := range propCtx {
		if mp, ok := inter.(map[string]interface{}); ok {
			result = append(result, mp)
		}
	}

	return result
}

func mergeCtx(local, global map[string]interface{}) map[string]interface{} {
	ctx := make(map[string]interface{})

	for k, v := range global {
		ctx[k] = v
	}

	for k, v := range local {
		ctx[k] = v
	}

	return ctx
}

func getRangeProp(in *xlsx.Row) string {
	if len(in.Cells) != 0 {
		match := rangeRgx.FindAllStringSubmatch(in.Cells[0].Value, -1)
		if match != nil {
			return match[0][1]
		}
	}

	return ""
}
