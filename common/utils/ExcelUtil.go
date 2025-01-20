package utils

import (
	"fmt"
	"github.com/mitchellh/mapstructure"
	"github.com/xuri/excelize/v2"
	"io"
	"log"
	"reflect"
)

type File interface {
	io.Reader
	io.ReaderAt
	io.Seeker
	io.Closer
}

// Style 样式
type Style struct {
	Border        []Border    `json:"border"`
	Fill          Fill        `json:"fill"`
	Font          *Font       `json:"font"`
	Alignment     *Alignment  `json:"alignment"`
	Protection    *Protection `json:"protection"`
	NumFmt        int         `json:"number_format"`
	DecimalPlaces int         `json:"decimal_places"`
	CustomNumFmt  *string     `json:"custom_number_format"`
	Lang          string      `json:"lang"`
	NegRed        bool        `json:"negred"`
}

// Border 边框
type Border struct {
	Type  string `json:"type"`
	Color string `json:"color"`
	Style int    `json:"style"`
}

// Fill 填充
type Fill struct {
	Type    string   `json:"type"`
	Pattern int      `json:"pattern"`
	Color   []string `json:"color"`
	Shading int      `json:"shading"`
}

// Font 字体
type Font struct {
	Bold      bool    `json:"bold"`      // 是否加粗
	Italic    bool    `json:"italic"`    // 是否倾斜
	Underline string  `json:"underline"` // single    double
	Family    string  `json:"family"`    // 字体样式
	Size      float64 `json:"size"`      // 字体大小
	Strike    bool    `json:"strike"`    // 删除线
	Color     string  `json:"color"`     // 字体颜色
}

// Protection 保护
type Protection struct {
	Hidden bool `json:"hidden"`
	Locked bool `json:"locked"`
}

// Alignment 对齐
type Alignment struct {
	Horizontal      string `json:"horizontal"`        // 水平对齐方式
	Indent          int    `json:"indent"`            // 缩进  只要设置了值，就变成了左对齐
	JustifyLastLine bool   `json:"justify_last_line"` // 两端分散对齐，只有在水平对齐选择 distributed 时起作用
	ReadingOrder    uint64 `json:"reading_order"`     // 文字方向 不知道值范围和具体的含义
	RelativeIndent  int    `json:"relative_indent"`   // 不知道具体的含义
	ShrinkToFit     bool   `json:"shrink_to_fit"`     // 缩小字体填充
	TextRotation    int    `json:"text_rotation"`     // 文本旋转
	Vertical        string `json:"vertical"`          // 垂直对齐
	WrapText        bool   `json:"wrap_text"`         // 自动换行
}

var cellCols = []string{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
	"L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}

func ReadExcel(file File, output interface{}) ([]interface{}, error) {
	f, err := excelize.OpenReader(file)
	if err != nil {
		log.Println(err.Error())
		return nil, err
	}

	firstSheet := f.GetSheetName(0)
	rows, err := f.GetRows(firstSheet)
	if err != nil {
		log.Println(err.Error())
		return nil, err
	}

	var title []string
	var ret []map[string]string
	for i, row := range rows {
		if i == 0 {
			for _, colCell := range row {
				title = append(title, colCell)
			}
		} else {
			theRow := map[string]string{}
			for j, colCell := range row {
				k := title[j]
				theRow[k] = colCell
			}
			ret = append(ret, theRow)
		}
	}

	slice := make([]interface{}, 0)
	for _, data := range ret {
		err = mapstructure.WeakDecode(data, &output)
		if err != nil {
			log.Println(err.Error())
			return nil, err
		}
		slice = append(slice, output)
	}
	return slice, nil
}

func WriteExcel(sheetName string, rType reflect.Type, d []map[string]interface{}) (*excelize.File, error) {
	// 创建File
	file := excelize.NewFile()

	// 创建Sheet
	if sheetName == "" {
		sheetName = "Sheet1"
	}
	file.NewSheet(sheetName)

	// 定义表头样式（通过结构体方式指定）
	headStyle, err := file.NewStyle(getHeaderStyle())
	if err != nil {
		log.Println(err.Error())
		return nil, err
	}

	// 定义行样式（通过JSON格式指定）
	rowStyle, err := file.NewStyle(getRowStyle())
	if err != nil {
		log.Println(err.Error())
		return nil, err
	}

	fieldNum := rType.NumField()
	tIndex := 0
	for i := 0; i < fieldNum; i++ {
		field := rType.Field(i)
		tag := field.Tag.Get("json")
		if tag == "" {
			continue
		}
		cel := fmt.Sprintf("%s%d", cellCols[tIndex], 1)
		// 写入表头内容
		_ = file.SetCellValue(sheetName, cel, tag)
		// 设置表头样式
		_ = file.SetCellStyle(sheetName, cel, cel, headStyle)
		// 下一列
		tIndex++
	}

	// 循环写入数据
	line := 1
	for _, value := range d {
		line++
		cIndex := 0
		for i := 0; i < fieldNum; i++ {
			field := rType.Field(i)
			tag := field.Tag.Get("json")
			if tag == "" {
				continue
			}
			cel := fmt.Sprintf("%s%d", cellCols[cIndex], line)
			// 写入行内容
			v, ok := value[field.Name].(string)
			if ok {
				_ = file.SetCellValue(sheetName, cel, v)
			} else {
				_ = file.SetCellValue(sheetName, cel, value[field.Name])
			}
			// 设置行样式
			_ = file.SetCellStyle(sheetName, cel, cel, rowStyle)
			// 下一列
			cIndex++
		}
	}
	return file, nil
}

func getHeaderStyle() *excelize.Style {
	return &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "right",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "#000000",
				Style: 1,
			},
		}, Fill: excelize.Fill{
			// gradient： 渐变色    pattern   填充图案
			// Pattern: 1,                   // 填充样式  当类型是 pattern 0-18 填充图案  1 实体填充
			// Color:   []string{"#FF0000"}, // 当Type = pattern 时，只有一个
			Type:  "gradient",
			Color: []string{"#DDEBF7", "##DDEBF7"},
			// 类型是 gradient 使用 0-5 横向(每种颜色横向分布) 纵向 对角向上 对角向下 有外向内 由内向外
			Shading: 1,
		}, Font: &excelize.Font{
			// Bold: true,    // 假哭
			// Italic: false,  // 斜体
			// Underline: "single",  // 下划线
			Size:   11,
			Family: "宋体",
			// Strike:    true, // 删除线
			// Color: "#00FFFF",
		}, Alignment: &excelize.Alignment{
			// 水平对齐方式 center left right fill(填充) justify(两端对齐)  centerContinuous(跨列居中) distributed(分散对齐)
			Horizontal: "left",
			// 垂直对齐方式 center top  justify distributed
			Vertical: "center",
			// Indent:     1,        // 缩进  只要有值就变成了左对齐 + 缩进
			// TextRotation: 30, // 旋转
			// RelativeIndent:  10,   // 好像没啥用
			// ReadingOrder:    0,    // 不知道怎么设置
			// JustifyLastLine: true, // 两端分散对齐，只有 水平对齐 为 distributed 时 设置true 才有效
			// WrapText:        true, // 自动换行
			// ShrinkToFit:     true, // 缩小字体以填充单元格
		}, Protection: &excelize.Protection{
			Hidden: true,
			Locked: true,
		},
		// 内置的数字格式样式   0-638  常用的 0-58  配合lang使用，因为语言不同样式不同 具体的样式参照文档
		NumFmt: 0,
		// zh-cn 中文
		//Lang: "zh-cn",
		// 小数位数  只有NumFmt是 2-11 有效
		// CustomNumFmt: "",// 自定义样式  是指针，只能通过变量的方式
		DecimalPlaces: new(int),
		NegRed:        true,
	}
}

func getRowStyle() *excelize.Style {
	return &excelize.Style{
		Border: []excelize.Border{
			{
				Type:  "right",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "left",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "top",
				Color: "#000000",
				Style: 1,
			},
			{
				Type:  "bottom",
				Color: "#000000",
				Style: 1,
			},
		}, Font: &excelize.Font{
			Size:   11,
			Family: "宋体",
		}, Alignment: &excelize.Alignment{
			Horizontal: "left",
			Vertical:   "center",
		},
	}
}
