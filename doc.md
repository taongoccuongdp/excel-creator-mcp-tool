# ExcelMCP User Guide

## Tổng quan

ExcelMCP là một MCP server cung cấp 1 tool duy nhất:

- `create_excel_file`

Tool này dùng để:

1. Nhận dữ liệu từ MCP client
2. Tạo file Excel `.xlsx`
3. Lưu file vào thư mục output trên server
4. Upload file sau khi tạo xong
5. Trả lại thông tin file và `download_url`

## Thông tin server

MCP server đang chạy tại:

- `http://host:port/mcp`

Các endpoint hữu ích:

- MCP endpoint: `POST http://host:port/mcp`
- SSE endpoint: `GET http://host:port/sse`
- Health check: `GET http://host:port/healthz`

Ví dụ health response:

```json
{"status":"ok","service":"excel-creator"}
```

## Cách dùng MCP tool

Sau khi kết nối MCP server, client chỉ cần gọi tool:

- `create_excel_file`

Tool này có 2 kiểu sử dụng chính:

1. Simple mode
2. Advanced mode với `workbook_json`

### 1. Simple mode

Dùng khi cần tạo một sheet dạng bảng đơn giản, có:

- header
- dữ liệu nhiều dòng
- độ rộng cột
- freeze header
- auto filter
- style cơ bản

Các field quan trọng:

| Field | Bắt buộc | Kiểu | Ý nghĩa |
| --- | --- | --- | --- |
| `output_path` | Có | `string` | Đường dẫn file Excel cần tạo, ví dụ `demo/report.xlsx` |
| `task` | Có | `string` | Tên task gửi kèm khi upload |
| `sheet_name` | Không | `string` | Tên sheet, mặc định là `Sheet1` |
| `overwrite` | Không | `boolean` | Ghi đè file cũ nếu đã tồn tại |
| `headers_json` | Không | `string` | JSON string của mảng header |
| `rows_json` | Không | `string` | JSON string của dữ liệu |
| `column_widths_json` | Không | `string` | JSON string object cấu hình độ rộng cột |
| `freeze_header` | Không | `boolean` | Freeze dòng đầu |
| `auto_filter` | Không | `boolean` | Bật filter cho header |
| `header_fill_color` | Không | `string` | Màu nền header |
| `header_font_color` | Không | `string` | Màu chữ header |
| `header_bold` | Không | `boolean` | In đậm header |
| `header_border` | Không | `boolean` | Viền header |
| `body_wrap_text` | Không | `boolean` | Wrap text cho body |
| `body_border` | Không | `boolean` | Viền body |
| `body_number_format` | Không | `string` | Format số cho body, ví dụ `#,##0` |

Ví dụ payload:

```json
{
  "output_path": "demo/price-report.xlsx",
  "task": "price_report",
  "sheet_name": "Prices",
  "headers_json": "[\"Product\",\"Price\",\"Quantity\"]",
  "rows_json": "[[\"Apple\",12000,5],[\"Orange\",9000,8],[\"Pear\",15000,3]]",
  "column_widths_json": "{\"Product\":24,\"B\":14,\"C\":12}",
  "freeze_header": true,
  "auto_filter": true,
  "header_fill_color": "#D9EAF7",
  "header_font_color": "#1F2937",
  "header_bold": true,
  "header_border": true,
  "body_border": true,
  "body_number_format": "#,##0"
}
```

Lưu ý:

- `headers_json`, `rows_json`, `column_widths_json` đều là `string`
- Bên trong các field này là JSON đã được stringify
- `rows_json` có thể là mảng các mảng hoặc mảng các object

Ví dụ `rows_json` dạng object:

```json
"[{\"Product\":\"Apple\",\"Price\":12000,\"Quantity\":5},{\"Product\":\"Orange\",\"Price\":9000,\"Quantity\":8}]"
```

### 2. Advanced mode với `workbook_json`

Dùng khi cần kiểm soát chi tiết hơn:

- ghi từng cell
- merge cell
- row height
- column width
- hyperlink
- freeze panes
- style theo từng cell

Các field quan trọng:

| Field | Bắt buộc | Kiểu | Ý nghĩa |
| --- | --- | --- | --- |
| `output_path` | Có | `string` | Đường dẫn file Excel cần tạo |
| `task` | Có | `string` | Tên task upload |
| `sheet_name` | Không | `string` | Tên sheet fallback |
| `overwrite` | Không | `boolean` | Ghi đè file cũ |
| `template_path` | Không | `string` | File template `.xlsx` trên server |
| `workbook_json` | Có trong mode này | `string` | JSON string của workbook spec |

Ví dụ payload:

```json
{
  "output_path": "demo/advanced-report.xlsx",
  "task": "advanced_report",
  "workbook_json": "{\"sheet_name\":\"Report\",\"column_widths\":[{\"from\":\"A\",\"to\":\"A\",\"width\":24},{\"from\":\"B\",\"to\":\"C\",\"width\":16}],\"merges\":[{\"start\":\"A1\",\"end\":\"C1\"}],\"freeze_panes\":{\"freeze\":true,\"y_split\":1,\"top_left_cell\":\"A2\",\"active_pane\":\"bottomLeft\"},\"styles\":{\"title\":{\"bold\":true,\"horizontal\":\"center\",\"vertical\":\"center\"},\"money\":{\"horizontal\":\"right\",\"number_format\":\"#,##0\"}},\"cells\":[{\"ref\":\"A1\",\"value\":\"Price Report\",\"style_ref\":\"title\"},{\"ref\":\"A2\",\"value\":\"Apple\"},{\"ref\":\"B2\",\"value\":3000000,\"style_ref\":\"money\"}]}"
}
```

### Cấu trúc chính trong `workbook_json`

| Field | Ý nghĩa |
| --- | --- |
| `sheet_name` | Tên sheet |
| `column_widths` | Danh sách cấu hình độ rộng cột |
| `row_heights` | Danh sách chiều cao dòng |
| `merges` | Danh sách vùng merge |
| `freeze_panes` | Cấu hình cố định pane |
| `styles` | Map tên style |
| `cells` | Danh sách cell cần ghi |

### `StyleSpec`

Các thuộc tính style đang hỗ trợ:

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `fill_color` | `string` | Màu nền |
| `font_color` | `string` | Màu chữ |
| `bold` | `boolean` | In đậm |
| `underline` | `boolean` | Gạch chân |
| `font_size` | `number` | Cỡ chữ |
| `horizontal` | `string` | Canh ngang |
| `vertical` | `string` | Canh dọc |
| `wrap_text` | `boolean` | Wrap text |
| `number_format` | `string` | Format số, ví dụ `#,##0`, `#,##0.00` |
| `border` | `boolean` | Viền 4 cạnh |
| `border_color` | `string` | Màu viền |

Ví dụ style tiền:

```json
{
  "body_money": {
    "border": true,
    "border_color": "#1F2937",
    "horizontal": "right",
    "vertical": "center",
    "number_format": "#,##0"
  }
}
```

Lưu ý về format số:

- Muốn format số hoạt động, `value` phải là số thật, ví dụ `3000000`
- Không nên truyền `"3000000"` dưới dạng string
- Dấu hiển thị `.` hay `,` phụ thuộc vào locale Excel của máy mở file
- Thông thường nên dùng `#,##0` hoặc `#,##0.00`

### Ghi text in đậm bằng HTML tag

Nếu `value` là string, tool hỗ trợ:

- `<b>...</b>`
- `<strong>...</strong>`
- `<br>`

Ví dụ:

```json
{
  "ref": "A1",
  "value": "Tổng tiền: <b>3.000.000</b> VND"
}
```

## Dùng template

Nếu muốn mở từ template Excel có sẵn, truyền thêm:

- `template_path`

Ví dụ:

```json
{
  "output_path": "reports/from-template.xlsx",
  "task": "template_report",
  "template_path": "base/report-template.xlsx",
  "sheet_name": "Summary",
  "headers_json": "[\"Name\",\"Revenue\"]",
  "rows_json": "[[\"Alice\",1200],[\"Bob\",950]]",
  "overwrite": true
}
```

## Output của tool

Sau khi tạo file thành công, tool trả object với các field:

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `output_path` | `string` | Đường dẫn tuyệt đối của file đã tạo trên server |
| `sheet_name` | `string` | Tên sheet đã ghi |
| `bytes_written` | `integer` | Kích thước file |
| `used_template` | `boolean` | Có dùng template hay không |
| `task` | `string` | Task đã gửi |
| `upload_message` | `string` | Message từ API upload |
| `upload_status` | `integer` | Trạng thái từ API upload |
| `download_url` | `string` | URL để tải file |

Tool trả structured output qua `structuredContent` của MCP. Khi gọi từ Dify, các field này thường xuất hiện như output variables riêng; object `text/files/json` trong UI là cách Dify render kết quả, không phải toàn bộ payload MCP gốc.

Ví dụ output:

```json
{
  "output_path": "/path/to/out/demo/price-report.xlsx",
  "sheet_name": "Prices",
  "bytes_written": 8421,
  "used_template": false,
  "task": "price_report",
  "upload_message": "upload success",
  "upload_status": 200,
  "download_url": "http://example.com/download/price-report.xlsx"
}
```

## Lưu ý quan trọng khi sử dụng

1. `output_path` là bắt buộc và phải kết thúc bằng `.xlsx`
2. `task` là bắt buộc
3. Nếu file đã tồn tại và không truyền `overwrite=true`, tool sẽ trả lỗi
4. Nếu dùng `workbook_json`, các field simple mode như `headers_json`, `rows_json` sẽ bị bỏ qua
5. Upload là bước bắt buộc trong flow hiện tại
6. Nếu upload lỗi, lệnh sẽ trả lỗi dù file local có thể đã được tạo thành công

## Tóm tắt nhanh

Nếu chỉ cần bảng đơn giản:

- dùng `headers_json` + `rows_json`

Nếu cần format từng ô, merge, style riêng, hyperlink:

- dùng `workbook_json`

Nếu cần định dạng tiền hoặc số:

- dùng `body_number_format` trong simple mode
- hoặc `number_format` trong `StyleSpec` ở advanced mode
