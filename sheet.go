package main

import (
	"errors"
	"fmt"
	"strings"

	"github.com/xuri/excelize/v2"
)

func openWorkbook(templatePath string, templateBase string) (*excelize.File, bool, error) {
	if strings.TrimSpace(templatePath) == "" {
		return excelize.NewFile(), false, nil
	}

	resolvedTemplate, err := resolvePathInsideBase(templateBase, templatePath, ".xlsx")
	if err != nil {
		return nil, false, fmt.Errorf("invalid template_path: %w", err)
	}

	f, err := excelize.OpenFile(resolvedTemplate)
	if err != nil {
		return nil, false, fmt.Errorf("open template: %w", err)
	}

	return f, true, nil
}

func ensureSheets(f *excelize.File, sheets []SheetSpec, usedTemplate bool) error {
	if len(sheets) == 0 {
		return nil
	}

	existing := make(map[string]bool)
	for _, name := range f.GetSheetList() {
		existing[name] = true
	}

	if !usedTemplate {
		first := strings.TrimSpace(sheets[0].Name)
		if first == "" {
			return errors.New("first sheet name is required")
		}
		defaultName := f.GetSheetName(0)
		if defaultName != first {
			if err := f.SetSheetName(defaultName, first); err != nil {
				return fmt.Errorf("rename default sheet: %w", err)
			}
			delete(existing, defaultName)
			existing[first] = true
		}
	}

	for _, s := range sheets {
		if strings.TrimSpace(s.Name) == "" {
			return errors.New("sheet name is required")
		}
		if !existing[s.Name] {
			if _, err := f.NewSheet(s.Name); err != nil {
				return fmt.Errorf("create sheet %q: %w", s.Name, err)
			}
			existing[s.Name] = true
		}
	}

	return nil
}

func applySheet(f *excelize.File, sheet SheetSpec, styles map[string]StyleSpec, styleCache map[string]int) error {
	if strings.TrimSpace(sheet.Name) == "" {
		return errors.New("sheet name is required")
	}

	for _, cw := range sheet.ColumnWidths {
		if err := f.SetColWidth(sheet.Name, cw.From, cw.To, cw.Width); err != nil {
			return fmt.Errorf("set column width %s:%s: %w", cw.From, cw.To, err)
		}
	}

	for _, rh := range sheet.RowHeights {
		if err := f.SetRowHeight(sheet.Name, rh.Row, rh.Height); err != nil {
			return fmt.Errorf("set row height %d: %w", rh.Row, err)
		}
	}

	for _, c := range sheet.Cells {
		if err := writeCell(f, sheet.Name, c, styles, styleCache); err != nil {
			return err
		}
	}

	for _, mg := range sheet.Merges {
		if err := f.MergeCell(sheet.Name, mg.Start, mg.End); err != nil {
			return fmt.Errorf("merge %s:%s: %w", mg.Start, mg.End, err)
		}
	}

	if sheet.FreezePanes != nil {
		if err := f.SetPanes(sheet.Name, toExcelizePanes(sheet.FreezePanes)); err != nil {
			return fmt.Errorf("set panes: %w", err)
		}
	}

	return nil
}

func writeCell(f *excelize.File, sheetName string, c CellSpec, styles map[string]StyleSpec, styleCache map[string]int) error {
	// display value of cell at default
	displayValue := c.Value

	// display the hyperlink if cell is an hyperlink type
	if c.Hyperlink != nil && c.Hyperlink.Display != "" {
		displayValue = c.Hyperlink.Display
	}

	// display the formula if cell is an formula
	if c.Formula != nil {
		displayValue = c.Formula
	}

	// set content value to cell (display value)
	if err := setCellContent(f, sheetName, c.Ref, displayValue); err != nil {
		return err
	}

	// set hyperlink to this cell if this cell contains a hyperlink
	if c.Hyperlink != nil && strings.TrimSpace(c.Hyperlink.URL) != "" {
		opts := make([]excelize.HyperlinkOpts, 0, 1)
		if c.Hyperlink.Display != "" || c.Hyperlink.Tooltip != "" {
			opt := excelize.HyperlinkOpts{}
			if c.Hyperlink.Display != "" {
				opt.Display = ptr(c.Hyperlink.Display)
			}
			if c.Hyperlink.Tooltip != "" {
				opt.Tooltip = ptr(c.Hyperlink.Tooltip)
			}
			opts = append(opts, opt)
		}
		if err := f.SetCellHyperLink(sheetName, c.Ref, c.Hyperlink.URL, "External", opts...); err != nil {
			return fmt.Errorf("set hyperlink %s: %w", c.Ref, err)
		}
	}

	// set formula to this cell if this cell contains a Formula
	if c.Formula != nil && strings.TrimSpace(c.Formula.Formula) != "" {
		opts := make([]excelize.FormulaOpts, 0, 1)
		if c.Formula.FormulaTypeArray != "" || c.Formula.FormulaTypeDataTable != "" || c.Formula.FormulaTypeShared != "" {
			opt := excelize.FormulaOpts{}
			if c.Formula.FormulaTypeArray != "" {
				opt.Type = ptr(c.Formula.FormulaTypeArray)
			}
			if c.Formula.FormulaTypeDataTable != "" {
				opt.Type = ptr(c.Formula.FormulaTypeDataTable)
			}
			if c.Formula.FormulaTypeShared != "" {
				opt.Type = ptr(c.Formula.FormulaTypeShared)
			}
			opt.Ref = ptr(c.Ref)
			opts = append(opts, opt)
		}
		if err := f.SetCellFormula(sheetName, c.Ref, c.Formula.Formula, opts...); err != nil {
			return fmt.Errorf("set formula %s: %w", c.Ref, err)
		}
	}

	if c.StyleRef != "" {
		styleID, err := resolveStyleID(f, styles, c.StyleRef, styleCache)
		if err != nil {
			return err
		}
		if err := f.SetCellStyle(sheetName, c.Ref, c.Ref, styleID); err != nil {
			return fmt.Errorf("style cell %s: %w", c.Ref, err)
		}
	}

	return nil
}

func setCellContent(f *excelize.File, sheetName string, cellRef string, value any) error {

	// try to write bold rich text
	runs, ok, err := parseHTMLBoldRichText(value)
	if err != nil {
		return fmt.Errorf("parse rich text %s: %w", cellRef, err)
	}
	if ok {
		if len(runs) == 0 {
			if err := f.SetCellValue(sheetName, cellRef, ""); err != nil {
				return fmt.Errorf("set empty rich text %s: %w", cellRef, err)
			}
			return nil
		}
		if err := f.SetCellRichText(sheetName, cellRef, runs); err != nil {
			return fmt.Errorf("set rich text %s: %w", cellRef, err)
		}
		return nil
	}

	// standard value
	if err := f.SetCellValue(sheetName, cellRef, value); err != nil {
		return fmt.Errorf("set value %s: %w", cellRef, err)
	}
	return nil
}
