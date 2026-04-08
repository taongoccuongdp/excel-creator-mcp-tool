package main

import (
	"strings"

	"github.com/xuri/excelize/v2"
)

func toExcelizeStyle(s StyleSpec) (*excelize.Style, error) {
	out := &excelize.Style{}

	if s.Border {
		color := normalizeColor(s.BorderColor, "000000")
		out.Border = []excelize.Border{
			{Type: "left", Color: color, Style: 1},
			{Type: "right", Color: color, Style: 1},
			{Type: "top", Color: color, Style: 1},
			{Type: "bottom", Color: color, Style: 1},
		}
	}

	if s.FillColor != "" {
		out.Fill = excelize.Fill{
			Type:    "pattern",
			Pattern: 1,
			Color:   []string{normalizeColor(s.FillColor, "")},
		}
	}

	if s.FontColor != "" || s.Bold || s.Underline || s.FontSize > 0 {
		out.Font = &excelize.Font{
			Bold:  s.Bold,
			Color: normalizeColor(s.FontColor, ""),
		}
		if s.Underline {
			out.Font.Underline = "single"
		}
		if s.FontSize > 0 {
			out.Font.Size = float64(s.FontSize)
		}
	}

	if s.WrapText || s.Horizontal != "" || s.Vertical != "" {
		out.Alignment = &excelize.Alignment{
			WrapText:   s.WrapText,
			Horizontal: s.Horizontal,
			Vertical:   s.Vertical,
		}
	}

	if format := strings.TrimSpace(s.NumberFormat); format != "" {
		out.CustomNumFmt = ptr(format)
	}

	return out, nil
}

func toExcelizePanes(in *FreezePanesSpec) *excelize.Panes {
	if in == nil {
		return nil
	}
	return &excelize.Panes{
		Freeze:      in.Freeze,
		XSplit:      in.XSplit,
		YSplit:      in.YSplit,
		TopLeftCell: in.TopLeftCell,
		ActivePane:  in.ActivePane,
	}
}

func normalizeColor(color string, fallback string) string {
	c := strings.TrimSpace(strings.TrimPrefix(color, "#"))
	if c == "" {
		return fallback
	}
	return strings.ToUpper(c)
}
