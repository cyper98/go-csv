package go_csv

import (
	"fmt"
	"strings"
)

type Font struct {
	Bold      bool   `xml:"b,omitempty"`
	Italic    bool   `xml:"i,omitempty"`
	Strike    bool   `xml:"strike,omitempty"`
	Underline string `xml:"u,omitempty"`
	Size      int    `xml:"sz>val"`
	Color     string `xml:"color>rgb,omitempty"`
	Family    string `xml:"family"`
	Name      string `xml:"name>val"`
}

type Fill struct {
	Pattern string `xml:"patternFill>patternType,omitempty"`
	FGColor string `xml:"patternFill>fgColor>rgb,omitempty"`
	BGColor string `xml:"patternFill>bgColor>rgb,omitempty"`
}

type Border struct {
	Left   string `xml:"left>style,omitempty"`
	Right  string `xml:"right>style,omitempty"`
	Top    string `xml:"top>style,omitempty"`
	Bottom string `xml:"bottom>style,omitempty"`
	Color  string `xml:"left>color>rgb,omitempty"`
}

type Alignment struct {
	Horizontal string `xml:"horizontal,omitempty"`
	Vertical   string `xml:"vertical,omitempty"`
	Wrap       bool   `xml:"wrapText,omitempty"`
	Rotation   int    `xml:"textRotation,omitempty"`
	Indent     int    `xml:"indent,omitempty"`
}

type Style struct {
	Font      *Font
	Fill      *Fill
	Border    *Border
	Alignment *Alignment
	NumFmt    int
	Hidden    bool
	Lock      bool
}

func (s *Style) ToXML() (string, error) {
	if s == nil {
		return "", nil
	}

	var parts []string

	if s.Font != nil {
		fontParts := []string{}
		if s.Font.Bold {
			fontParts = append(fontParts, `<b/>`)
		}
		if s.Font.Italic {
			fontParts = append(fontParts, `<i/>`)
		}
		if s.Font.Strike {
			fontParts = append(fontParts, `<strike/>`)
		}
		if s.Font.Underline != "" {
			fontParts = append(fontParts, fmt.Sprintf(`<u val="%s"/>`, s.Font.Underline))
		}
		if s.Font.Size > 0 {
			fontParts = append(fontParts, fmt.Sprintf(`<sz val="%d"/>`, s.Font.Size))
		}
		if s.Font.Color != "" {
			fontParts = append(fontParts, fmt.Sprintf(`<color rgb="%s"/>`, s.Font.Color))
		}
		if s.Font.Name != "" {
			fontParts = append(fontParts, fmt.Sprintf(`<name val="%s"/>`, s.Font.Name))
		}
		if len(fontParts) > 0 {
			parts = append(parts, fmt.Sprintf(`<font>%s</font>`, strings.Join(fontParts, "")))
		}
	}

	if s.Fill != nil {
		fill := fmt.Sprintf(`<fill><patternFill patternType="%s"`, s.Fill.Pattern)
		if s.Fill.FGColor != "" {
			fill += fmt.Sprintf(`><fgColor rgb="%s"/></patternFill>`, s.Fill.FGColor)
		} else {
			fill += "/>"
		}
		parts = append(parts, fill+"</fill>")
	}

	if s.Border != nil {
		border := "<border>"
		if s.Border.Left != "" {
			border += fmt.Sprintf(`<left style="%s"><color rgb="%s"/></left>`, s.Border.Left, s.Border.Color)
		} else {
			border += "<left/>"
		}
		if s.Border.Right != "" {
			border += fmt.Sprintf(`<right style="%s"><color rgb="%s"/></right>`, s.Border.Right, s.Border.Color)
		} else {
			border += "<right/>"
		}
		if s.Border.Top != "" {
			border += fmt.Sprintf(`<top style="%s"><color rgb="%s"/></top>`, s.Border.Top, s.Border.Color)
		} else {
			border += "<top/>"
		}
		if s.Border.Bottom != "" {
			border += fmt.Sprintf(`<bottom style="%s"><color rgb="%s"/></bottom>`, s.Border.Bottom, s.Border.Color)
		} else {
			border += "<bottom/>"
		}
		border += "</border>"
		parts = append(parts, border)
	}

	if s.Alignment != nil {
		align := `<alignment`
		if s.Alignment.Horizontal != "" {
			align += fmt.Sprintf(` horizontal="%s"`, s.Alignment.Horizontal)
		}
		if s.Alignment.Vertical != "" {
			align += fmt.Sprintf(` vertical="%s"`, s.Alignment.Vertical)
		}
		if s.Alignment.Wrap {
			align += ` wrapText="1"`
		}
		if s.Alignment.Indent > 0 {
			align += fmt.Sprintf(` indent="%d"`, s.Alignment.Indent)
		}
		align += "/>"
		parts = append(parts, align)
	}

	if len(parts) == 0 {
		return "", nil
	}

	return strings.Join(parts, ""), nil
}

type RichTextRun struct {
	Text   string `xml:"t"`
	Bold   bool   `xml:"rPr>b,omitempty"`
	Italic bool   `xml:"rPr>i,omitempty"`
	Size   int    `xml:"rPr>sz>val,omitempty"`
	Color  string `xml:"rPr>color>rgb,omitempty"`
	Font   string `xml:"rPr>rFont>val,omitempty"`
}

type RichText struct {
	Runs []RichTextRun `xml:"r"`
}

func (r *RichText) ToXML() string {
	if len(r.Runs) == 0 {
		return ""
	}

	var parts []string
	for _, run := range r.Runs {
		parts = append(parts, fmt.Sprintf(`<r>%s%s</r>`, r.runProps(run), run.Text))
	}

	return `<is>` + strings.Join(parts, "") + `</is>`
}

func (r *RichText) runProps(run RichTextRun) string {
	if run.Text == "" && !run.Bold && !run.Italic && run.Size == 0 && run.Color == "" && run.Font == "" {
		return ""
	}

	props := "<rPr>"
	if run.Bold {
		props += "<b/>"
	}
	if run.Italic {
		props += "<i/>"
	}
	if run.Size > 0 {
		props += fmt.Sprintf("<sz val=\"%d\"/>", run.Size)
	}
	if run.Color != "" {
		props += fmt.Sprintf("<color rgb=\"%s\"/>", run.Color)
	}
	if run.Font != "" {
		props += fmt.Sprintf("<rFont val=\"%s\"/>", run.Font)
	}
	props += "</rPr>"
	return props
}

type Hyperlink struct {
	Display string
	Tooltip string
	URL     string
}

type Comment struct {
	Author string
	Text   string
	Ref    string
}

type DataValidation struct {
	Type       string
	Operator   string
	Formula1   string
	Formula2   string
	Sqref      string
	ShowError  bool
	ErrorTitle string
	Error      string
	ShowInput  bool
	InputTitle string
	Input      string
}

func (dv *DataValidation) ToXML() string {
	xml := `<dataValidation type="%s" operator="%s" formula1="%s" sqref="%s" showErrorBox="%t"`
	xml += fmt.Sprintf(xml, dv.Type, dv.Operator, dv.Formula1, dv.Sqref, dv.ShowError)

	if dv.ErrorTitle != "" || dv.Error != "" {
		xml += fmt.Sprintf(`<error title="%s">%s</error>`, dv.ErrorTitle, dv.Error)
	}
	if dv.InputTitle != "" || dv.Input != "" {
		xml += fmt.Sprintf(`<prompt title="%s">%s</prompt>`, dv.InputTitle, dv.Input)
	}

	xml += "</dataValidation>"
	return xml
}

type Table struct {
	Name       string
	Range      string
	StyleName  string
	ShowHeader bool
	ShowTotals bool
	AutoFilter bool
}

type ChartType string

const (
	ChartLine          ChartType = "line"
	ChartLine3D        ChartType = "line3D"
	ChartColumn        ChartType = "col"
	ChartColumn3D      ChartType = "col3D"
	ChartColumnStacked ChartType = "colStacked"
	ChartBar           ChartType = "bar"
	ChartBar3D         ChartType = "bar3D"
	ChartBarStacked    ChartType = "barStacked"
	ChartArea          ChartType = "area"
	ChartArea3D        ChartType = "area3D"
	ChartPie           ChartType = "pie"
	ChartPie3D         ChartType = "pie3D"
	ChartPieOfPie      ChartType = "pieOfPie"
	ChartDoughnut      ChartType = "doughnut"
	ChartScatter       ChartType = "scatter"
	ChartRadar         ChartType = "radar"
	ChartRadarArea     ChartType = "radarArea"
	ChartSurface       ChartType = "surface"
	ChartStock         ChartType = "stock"
)

type ChartGrouping string

const (
	ChartGroupingStandard       ChartGrouping = "standard"
	ChartGroupingStacked        ChartGrouping = "stacked"
	ChartGroupingPercentStacked ChartGrouping = "percentStacked"
	ChartGroupingClustered      ChartGrouping = "clustered"
)

type ChartLegendPosition string

const (
	ChartLegendPositionRight  ChartLegendPosition = "r"
	ChartLegendPositionLeft   ChartLegendPosition = "l"
	ChartLegendPositionTop    ChartLegendPosition = "t"
	ChartLegendPositionBottom ChartLegendPosition = "b"
)

type ChartMarker struct {
	Type      string
	Size      int
	ForeColor string
	BackColor string
}

type ChartLineFormat struct {
	Color  string
	Width  float64
	Style  string
	Marker *ChartMarker
}

type ChartSeries struct {
	Name       string
	Categories string
	Values     string
	LineWidth  float64
	LineColor  string
	Marker     string
	Smooth     bool
	InvertFill bool
	LineFormat *ChartLineFormat
}

type Chart struct {
	Type        ChartType
	Title       string
	Series      []ChartSeries
	Height      int
	Width       int
	Legend      bool
	LegendPos   string
	XAxisLabel  string
	YAxisLabel  string
	Grouping    string
	Direction   string
	GapWidth    int
	GapDepth    int
	PlotOnly    bool
	VaryColors  bool
	ChartColors []string
}

func (c *Chart) ToXML() string {
	if c == nil {
		return ""
	}

	var parts []string
	parts = append(parts, fmt.Sprintf(`<chart type="%s">`, c.Type))

	if c.Title != "" {
		parts = append(parts, fmt.Sprintf(`<title><val>%s</val></title>`, xmlEscape(c.Title)))
	}

	if len(c.Series) > 0 {
		parts = append(parts, "<series>")
		for _, s := range c.Series {
			parts = append(parts, "<ser>")
			if s.Name != "" {
				parts = append(parts, fmt.Sprintf(`<tx><v>%s</v></tx>`, xmlEscape(s.Name)))
			}
			parts = append(parts, fmt.Sprintf(`<cat><val>%s</val></cat>`, xmlEscape(s.Categories)))
			parts = append(parts, fmt.Sprintf(`<val><numRef><f>%s</f></numRef></val>`, s.Values))
			parts = append(parts, "</ser>")
		}
		parts = append(parts, "</series>")
	}

	if c.XAxisLabel != "" {
		parts = append(parts, fmt.Sprintf(`<catAx><tickLblSink><val>%s</val></tickLblSink></catAx>`, c.XAxisLabel))
	}

	if c.Legend {
		legendPos := "r"
		if c.LegendPos != "" {
			legendPos = c.LegendPos
		}
		parts = append(parts, fmt.Sprintf(`<legend><legendPos val="%s"/></legend>`, legendPos))
	}

	parts = append(parts, "</chart>")

	return strings.Join(parts, "")
}

type SheetOpts struct {
	FreezePane  string
	Selection   []string
	TabColor    string
	Zoom        float64
	View        string
	ShowGrid    bool
	ShowHeaders bool
}

type RowOpts struct {
	Height float64
	Hidden bool
	Style  int
}

type ColOpts struct {
	Width  float64
	Hidden bool
	Style  int
}
