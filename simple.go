package main

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"github.com/xuri/excelize/v2"
)

func createExcelFile(ctx context.Context, _ *mcp.CallToolRequest, in CreateExcelInput) (*mcp.CallToolResult, CreateExcelOutput, error) {
	_ = ctx

	if strings.TrimSpace(in.Task) == "" {
		return nil, CreateExcelOutput{}, errors.New("task is required")
	}

	outputBase := envOrDefault("EXCEL_OUTPUT_DIR", "./out")
	templateBase := envOrDefault("EXCEL_TEMPLATE_DIR", "./templates")

	outputPath, err := resolvePathInsideBase(outputBase, in.OutputPath, ".xlsx")
	if err != nil {
		return nil, CreateExcelOutput{}, err
	}

	if err := os.MkdirAll(filepath.Dir(outputPath), 0o755); err != nil {
		return nil, CreateExcelOutput{}, fmt.Errorf("create output directory: %w", err)
	}

	if !in.Overwrite {
		if _, err := os.Stat(outputPath); err == nil {
			return nil, CreateExcelOutput{}, fmt.Errorf("file already exists: %s", outputPath)
		}
	}

	f, usedTemplate, err := openWorkbook(in.TemplatePath, templateBase)
	if err != nil {
		return nil, CreateExcelOutput{}, err
	}
	defer func() { _ = f.Close() }()

	sheet, styles, sheetName, autoFilterRange, err := buildWorkbook(in)
	if err != nil {
		return nil, CreateExcelOutput{}, err
	}

	if err := ensureSheets(f, []SheetSpec{sheet}, usedTemplate); err != nil {
		return nil, CreateExcelOutput{}, err
	}
	if err := applySheet(f, sheet, styles, make(map[string]int)); err != nil {
		return nil, CreateExcelOutput{}, err
	}

	if autoFilterRange != "" {
		if err := f.AutoFilter(sheetName, autoFilterRange, nil); err != nil {
			return nil, CreateExcelOutput{}, fmt.Errorf("apply autofilter: %w", err)
		}
	}

	if err := f.SaveAs(outputPath); err != nil {
		return nil, CreateExcelOutput{}, fmt.Errorf("save workbook: %w", err)
	}

	info, err := os.Stat(outputPath)
	if err != nil {
		return nil, CreateExcelOutput{}, fmt.Errorf("stat output file: %w", err)
	}

	uploadResult, err := uploadExcelFile(outputPath, in.Task)
	if err != nil {
		return nil, CreateExcelOutput{}, fmt.Errorf("upload workbook: %w", err)
	}

	out := CreateExcelOutput{
		OutputPath:    outputPath,
		SheetName:     sheetName,
		BytesWritten:  info.Size(),
		UsedTemplate:  usedTemplate,
		Task:          in.Task,
		UploadMessage: uploadResult.Message,
		UploadStatus:  uploadResult.Status,
		DownloadURL:   uploadResult.URL,
	}

	// Leave Content empty so go-sdk serializes the typed output into both
	// structuredContent and JSON text content. Dify can then expose the output
	// object consistently instead of only showing a plain success string.
	return &mcp.CallToolResult{}, out, nil
}

func buildInputSchema() json.RawMessage {
	s, _ := json.Marshal(map[string]any{
		"type": "object",
		"properties": map[string]any{
			"output_path": map[string]any{
				"type":        "string",
				"description": "Output file path relative to EXCEL_OUTPUT_DIR, for example reports/price.xlsx",
			},
			"task": map[string]any{
				"type":        "string",
				"description": "Upload task name, for example price_report.",
			},
			"sheet_name": map[string]any{
				"type":        "string",
				"description": "Worksheet name. Default is Sheet1.",
			},
			"overwrite": map[string]any{
				"type":        "boolean",
				"description": "Overwrite existing file if true.",
			},
			"template_path": map[string]any{
				"type":        "string",
				"description": "Optional template file path relative to EXCEL_TEMPLATE_DIR.",
			},
			"workbook_json": map[string]any{
				"type":        "string",
				"description": "Optional advanced workbook spec as JSON string. Supports styled cells, merges, row heights, hyperlinks, and freeze panes.",
			},
			"headers_json": map[string]any{
				"type":        "string",
				"description": "Optional JSON array of column headers, for example [\"Name\",\"Revenue\"].",
			},
			"rows_json": map[string]any{
				"type":        "string",
				"description": "JSON array of rows. Each row can be an array or an object. Example: [[\"Alice\",1200],[\"Bob\",950]] or [{\"Name\":\"Alice\",\"Revenue\":1200}].",
			},
			"column_widths_json": map[string]any{
				"type":        "string",
				"description": "Optional JSON object of column widths keyed by Excel column letter or header text, for example {\"A\":18,\"Revenue\":14}.",
			},
			"freeze_header": map[string]any{
				"type":        "boolean",
				"description": "Freeze the first row if headers exist.",
			},
			"auto_filter": map[string]any{
				"type":        "boolean",
				"description": "Enable Excel autofilter on the header row.",
			},
			"header_fill_color": map[string]any{
				"type":        "string",
				"description": "Header background color like #D9EAF7.",
			},
			"header_font_color": map[string]any{
				"type":        "string",
				"description": "Header font color like #1F2937.",
			},
			"header_bold": map[string]any{
				"type":        "boolean",
				"description": "Make header text bold.",
			},
			"header_border": map[string]any{
				"type":        "boolean",
				"description": "Add thin borders to header cells.",
			},
			"body_wrap_text": map[string]any{
				"type":        "boolean",
				"description": "Wrap body cell text.",
			},
			"body_border": map[string]any{
				"type":        "boolean",
				"description": "Add thin borders to body cells.",
			},
			"body_number_format": map[string]any{
				"type":        "string",
				"description": "Optional Excel custom number format for body cells, for example #,##0 or #,##0.00.",
			},
		},
		"required":             []string{"output_path", "task"},
		"additionalProperties": false,
	})
	return s
}

func buildOutputSchema() json.RawMessage {
	s, _ := json.Marshal(map[string]any{
		"type": "object",
		"properties": map[string]any{
			"output_path":    map[string]any{"type": "string"},
			"sheet_name":     map[string]any{"type": "string"},
			"bytes_written":  map[string]any{"type": "integer"},
			"used_template":  map[string]any{"type": "boolean"},
			"task":           map[string]any{"type": "string"},
			"upload_message": map[string]any{"type": "string"},
			"upload_status":  map[string]any{"type": "integer"},
			"download_url":   map[string]any{"type": "string"},
		},
		"required": []string{
			"output_path",
			"sheet_name",
			"bytes_written",
			"used_template",
			"task",
			"upload_message",
			"upload_status",
			"download_url",
		},
		"additionalProperties": false,
	})
	return s
}

func buildStyles(in CreateExcelInput) map[string]StyleSpec {
	styles := make(map[string]StyleSpec)

	header := StyleSpec{
		FillColor:   in.HeaderFillColor,
		FontColor:   in.HeaderFontColor,
		Bold:        in.HeaderBold,
		Border:      in.HeaderBorder,
		BorderColor: "D1D5DB",
	}
	if header != (StyleSpec{}) {
		styles["header"] = header
	}

	body := StyleSpec{
		WrapText:     in.BodyWrapText,
		NumberFormat: in.BodyNumberFormat,
		Border:       in.BodyBorder,
		BorderColor:  "E5E7EB",
	}
	if body != (StyleSpec{}) {
		styles["body"] = body
	}

	return styles
}

func buildWorkbook(in CreateExcelInput) (SheetSpec, map[string]StyleSpec, string, string, error) {
	if strings.TrimSpace(in.WorkbookJSON) != "" {
		spec, err := parseWorkbookJSON(in.WorkbookJSON, in.SheetName)
		if err != nil {
			return SheetSpec{}, nil, "", "", err
		}
		sheet := SheetSpec{
			Name:         spec.SheetName,
			ColumnWidths: spec.ColumnWidths,
			RowHeights:   spec.RowHeights,
			Merges:       spec.Merges,
			Cells:        spec.Cells,
			FreezePanes:  spec.FreezePanes,
		}
		return sheet, spec.Styles, spec.SheetName, "", nil
	}

	sheetName := strings.TrimSpace(in.SheetName)
	if sheetName == "" {
		sheetName = "Sheet1"
	}

	headers, err := parseHeadersJSON(in.HeadersJSON)
	if err != nil {
		return SheetSpec{}, nil, "", "", err
	}

	rows, derivedHeaders, err := parseRowsJSON(in.RowsJSON, headers)
	if err != nil {
		return SheetSpec{}, nil, "", "", err
	}
	if len(headers) == 0 {
		headers = derivedHeaders
	}
	if len(headers) == 0 && len(rows) == 0 {
		return SheetSpec{}, nil, "", "", errors.New("at least one of workbook_json or headers_json/rows_json is required")
	}

	columnWidths, err := parseColumnWidthsJSON(in.ColumnWidthsJSON, headers)
	if err != nil {
		return SheetSpec{}, nil, "", "", err
	}

	sheet := buildSheetSpec(sheetName, headers, rows, columnWidths, in)
	autoFilterRange := ""
	if in.AutoFilter && len(headers) > 0 {
		endCell, err := excelize.CoordinatesToCellName(max(1, len(headers)), max(1, len(rows)+1))
		if err != nil {
			return SheetSpec{}, nil, "", "", fmt.Errorf("compute autofilter range: %w", err)
		}
		autoFilterRange = "A1:" + endCell
	}
	return sheet, buildStyles(in), sheetName, autoFilterRange, nil
}

func parseWorkbookJSON(raw string, fallbackSheetName string) (*WorkbookSpec, error) {
	var spec WorkbookSpec
	if err := json.Unmarshal([]byte(raw), &spec); err != nil {
		return nil, fmt.Errorf("parse workbook_json: %w", err)
	}

	spec.SheetName = strings.TrimSpace(spec.SheetName)
	if spec.SheetName == "" {
		spec.SheetName = strings.TrimSpace(fallbackSheetName)
	}
	if spec.SheetName == "" {
		spec.SheetName = "Sheet1"
	}
	if len(spec.Cells) == 0 {
		return nil, errors.New("workbook_json must contain at least one cell")
	}
	if spec.Styles == nil {
		spec.Styles = map[string]StyleSpec{}
	}
	return &spec, nil
}

func buildSheetSpec(sheetName string, headers []string, rows [][]any, columnWidths []ColumnWidthSpec, in CreateExcelInput) SheetSpec {
	sheet := SheetSpec{
		Name:         sheetName,
		ColumnWidths: columnWidths,
	}

	if in.FreezeHeader && len(headers) > 0 {
		sheet.FreezePanes = &FreezePanesSpec{
			Freeze:      true,
			YSplit:      1,
			TopLeftCell: "A2",
			ActivePane:  "bottomLeft",
		}
	}

	if len(headers) > 0 {
		for i, header := range headers {
			ref, _ := excelize.CoordinatesToCellName(i+1, 1)
			cell := CellSpec{Ref: ref, Value: header}
			if hasHeaderStyle(in) {
				cell.StyleRef = "header"
			}
			sheet.Cells = append(sheet.Cells, cell)
		}
	}

	startRow := 1
	if len(headers) > 0 {
		startRow = 2
	}
	for rowIndex, row := range rows {
		for colIndex, value := range row {
			ref, _ := excelize.CoordinatesToCellName(colIndex+1, startRow+rowIndex)
			cell := CellSpec{Ref: ref, Value: value}
			if hasBodyStyle(in) {
				cell.StyleRef = "body"
			}
			sheet.Cells = append(sheet.Cells, cell)
		}
	}

	return sheet
}

func hasHeaderStyle(in CreateExcelInput) bool {
	return in.HeaderFillColor != "" || in.HeaderFontColor != "" || in.HeaderBold || in.HeaderBorder
}

func hasBodyStyle(in CreateExcelInput) bool {
	return in.BodyWrapText || in.BodyBorder || strings.TrimSpace(in.BodyNumberFormat) != ""
}

func parseHeadersJSON(raw string) ([]string, error) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return nil, nil
	}
	var headers []string
	if err := json.Unmarshal([]byte(raw), &headers); err != nil {
		return nil, fmt.Errorf("parse headers_json: %w", err)
	}
	return headers, nil
}

func parseRowsJSON(raw string, headers []string) ([][]any, []string, error) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return nil, nil, nil
	}

	var decoded []any
	if err := json.Unmarshal([]byte(raw), &decoded); err != nil {
		return nil, nil, fmt.Errorf("parse rows_json: %w", err)
	}
	if len(decoded) == 0 {
		return nil, headers, nil
	}

	switch first := decoded[0].(type) {
	case []any:
		rows := make([][]any, 0, len(decoded))
		for i, item := range decoded {
			row, ok := item.([]any)
			if !ok {
				return nil, nil, fmt.Errorf("rows_json row %d must be an array", i)
			}
			rows = append(rows, row)
		}
		return rows, headers, nil
	case map[string]any:
		derivedHeaders := headers
		if len(derivedHeaders) == 0 {
			derivedHeaders = make([]string, 0, len(first))
			for key := range first {
				derivedHeaders = append(derivedHeaders, key)
			}
			sort.Strings(derivedHeaders)
		}

		rows := make([][]any, 0, len(decoded))
		for i, item := range decoded {
			obj, ok := item.(map[string]any)
			if !ok {
				return nil, nil, fmt.Errorf("rows_json row %d must be an object", i)
			}
			row := make([]any, 0, len(derivedHeaders))
			for _, header := range derivedHeaders {
				row = append(row, obj[header])
			}
			rows = append(rows, row)
		}
		return rows, derivedHeaders, nil
	default:
		return nil, nil, errors.New("rows_json must be a JSON array of arrays or a JSON array of objects")
	}
}

func parseColumnWidthsJSON(raw string, headers []string) ([]ColumnWidthSpec, error) {
	raw = strings.TrimSpace(raw)
	if raw == "" {
		return nil, nil
	}

	var widths map[string]float64
	if err := json.Unmarshal([]byte(raw), &widths); err != nil {
		return nil, fmt.Errorf("parse column_widths_json: %w", err)
	}

	var specs []ColumnWidthSpec
	for key, width := range widths {
		column := strings.ToUpper(strings.TrimSpace(key))
		if column == "" {
			continue
		}
		if !isExcelColumn(column) {
			idx := indexOf(headers, key)
			if idx < 0 {
				return nil, fmt.Errorf("unknown column width key %q: use Excel column letters or a header value", key)
			}
			column, _ = excelize.ColumnNumberToName(idx + 1)
		}
		specs = append(specs, ColumnWidthSpec{From: column, To: column, Width: width})
	}

	sort.Slice(specs, func(i, j int) bool {
		return specs[i].From < specs[j].From
	})
	return specs, nil
}

func isExcelColumn(s string) bool {
	if s == "" {
		return false
	}
	for _, r := range s {
		if r < 'A' || r > 'Z' {
			return false
		}
	}
	return true
}

func indexOf(values []string, want string) int {
	for i, value := range values {
		if value == want {
			return i
		}
	}
	return -1
}
