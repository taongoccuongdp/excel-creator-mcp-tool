package main

type CreateExcelInput struct {
	OutputPath       string `json:"output_path"`
	Task             string `json:"task"`
	WorkbookJSON     string `json:"workbook_json,omitempty"`
	SheetName        string `json:"sheet_name,omitempty"`
	Overwrite        bool   `json:"overwrite,omitempty"`
	TemplatePath     string `json:"template_path,omitempty"`
	HeadersJSON      string `json:"headers_json,omitempty"`
	RowsJSON         string `json:"rows_json,omitempty"`
	ColumnWidthsJSON string `json:"column_widths_json,omitempty"`
	FreezeHeader     bool   `json:"freeze_header,omitempty"`
	AutoFilter       bool   `json:"auto_filter,omitempty"`
	HeaderFillColor  string `json:"header_fill_color,omitempty"`
	HeaderFontColor  string `json:"header_font_color,omitempty"`
	HeaderBold       bool   `json:"header_bold,omitempty"`
	HeaderBorder     bool   `json:"header_border,omitempty"`
	BodyWrapText     bool   `json:"body_wrap_text,omitempty"`
	BodyBorder       bool   `json:"body_border,omitempty"`
	BodyNumberFormat string `json:"body_number_format,omitempty"`
}

type CreateExcelOutput struct {
	OutputPath    string `json:"output_path"`
	SheetName     string `json:"sheet_name"`
	BytesWritten  int64  `json:"bytes_written"`
	UsedTemplate  bool   `json:"used_template"`
	Task          string `json:"task"`
	UploadMessage string `json:"upload_message"`
	UploadStatus  int    `json:"upload_status"`
	DownloadURL   string `json:"download_url"`
}

type SheetSpec struct {
	Name         string            `json:"name"`
	ColumnWidths []ColumnWidthSpec `json:"column_widths,omitempty"`
	RowHeights   []RowHeightSpec   `json:"row_heights,omitempty"`
	Merges       []MergeSpec       `json:"merges,omitempty"`
	Cells        []CellSpec        `json:"cells,omitempty"`
	FreezePanes  *FreezePanesSpec  `json:"freeze_panes,omitempty"`
}

type ColumnWidthSpec struct {
	From  string  `json:"from"`
	To    string  `json:"to"`
	Width float64 `json:"width"`
}

type RowHeightSpec struct {
	Row    int     `json:"row"`
	Height float64 `json:"height"`
}

type MergeSpec struct {
	Start string `json:"start"`
	End   string `json:"end"`
}

type CellSpec struct {
	Ref       string         `json:"ref"`
	Value     any            `json:"value,omitempty"`
	StyleRef  string         `json:"style_ref,omitempty"`
	Hyperlink *HyperlinkSpec `json:"hyperlink,omitempty"`
}

type HyperlinkSpec struct {
	URL     string `json:"url"`
	Display string `json:"display,omitempty"`
	Tooltip string `json:"tooltip,omitempty"`
}

type FreezePanesSpec struct {
	Freeze      bool   `json:"freeze,omitempty"`
	XSplit      int    `json:"x_split,omitempty"`
	YSplit      int    `json:"y_split,omitempty"`
	TopLeftCell string `json:"top_left_cell,omitempty"`
	ActivePane  string `json:"active_pane,omitempty"`
}

type StyleSpec struct {
	FillColor    string `json:"fill_color,omitempty"`
	FontColor    string `json:"font_color,omitempty"`
	Bold         bool   `json:"bold,omitempty"`
	Underline    bool   `json:"underline,omitempty"`
	FontSize     int    `json:"font_size,omitempty"`
	Horizontal   string `json:"horizontal,omitempty"`
	Vertical     string `json:"vertical,omitempty"`
	WrapText     bool   `json:"wrap_text,omitempty"`
	NumberFormat string `json:"number_format,omitempty"`
	Border       bool   `json:"border,omitempty"`
	BorderColor  string `json:"border_color,omitempty"`
}

type WorkbookSpec struct {
	SheetName    string               `json:"sheet_name"`
	ColumnWidths []ColumnWidthSpec    `json:"column_widths,omitempty"`
	RowHeights   []RowHeightSpec      `json:"row_heights,omitempty"`
	Merges       []MergeSpec          `json:"merges,omitempty"`
	Cells        []CellSpec           `json:"cells,omitempty"`
	FreezePanes  *FreezePanesSpec     `json:"freeze_panes,omitempty"`
	Styles       map[string]StyleSpec `json:"styles,omitempty"`
}

type UploadResponse struct {
	Message string `json:"message"`
	Status  int    `json:"status"`
	URL     string `json:"url"`
}
