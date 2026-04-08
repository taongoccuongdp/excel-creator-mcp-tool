# ExcelMCP

`ExcelMCP` là một MCP server viết bằng Go, cung cấp duy nhất một tool: `create_excel_file`.

Tool này nhận dữ liệu đầu vào, tạo file Excel `.xlsx`, lưu file vào thư mục output an toàn, upload file lên một endpoint HTTP, rồi trả về đường dẫn tải xuống cho client MCP.

Trọng tâm của project này là:

- Tạo Excel nhanh từ input đơn giản (`headers_json`, `rows_json`)
- Hỗ trợ cấu hình nâng cao qua `workbook_json`
- Kiểm soát an toàn đường dẫn khi đọc template và ghi file
- Trả về kết quả sẵn dùng cho workflow tự động hóa

Lưu ý quan trọng:

- Mỗi lần gọi tool hiện chỉ build một sheet.
- Upload là bước bắt buộc trong flow hiện tại.
- Nếu upload thất bại, lệnh sẽ trả lỗi dù file local có thể đã được lưu thành công.

## Chạy server

### Build và run

```bash
go build ./...
go vet ./...
go run .
MCP_ADDR=:3000 go run .
MCP_STDIO=1 go run .
go mod tidy
```

### Chế độ chạy

- HTTP mode là mặc định, lắng nghe ở `MCP_ADDR` hoặc `:8080`
- Stdio mode được bật khi `MCP_STDIO=1`

### HTTP endpoints

Khi chạy ở HTTP mode:

- `POST /mcp`: Streamable HTTP MCP endpoint
- `GET /sse`: SSE transport endpoint
- `GET /healthz`: health check
- `GET /`: trả về health check JSON

Ví dụ response health:

```json
{"status":"ok","service":"excel-creator"}
```

## Biến môi trường

| Biến | Mặc định | Ý nghĩa |
| --- | --- | --- |
| `MCP_ADDR` | `:8080` | Địa chỉ HTTP server |
| `MCP_STDIO` | rỗng | Nếu bằng `1` thì chạy ở stdio mode |
| `EXCEL_OUTPUT_DIR` | `./out` | Thư mục gốc cho file Excel đầu ra |
| `EXCEL_TEMPLATE_DIR` | `./templates` | Thư mục gốc cho file template `.xlsx` |
| `EXCEL_UPLOAD_URL` | `http://link/upload_genai/upload/` | Endpoint upload file sau khi tạo xong |

## Tool public: `create_excel_file`

### Chức năng

Tool `create_excel_file`:

1. Nhận input từ MCP client
2. Tạo workbook mới hoặc mở template
3. Build dữ liệu cho một sheet
4. Ghi sheet vào file `.xlsx`
5. Upload file qua HTTP multipart
6. Trả metadata và `download_url`

### Input bắt buộc

| Field | Kiểu | Bắt buộc | Ý nghĩa |
| --- | --- | --- | --- |
| `output_path` | `string` | Có | Đường dẫn file `.xlsx` cần tạo |
| `task` | `string` | Có | Giá trị gửi lên API upload trong form field `task` |

### Input tùy chọn

| Field | Kiểu | Dùng trong mode | Ý nghĩa |
| --- | --- | --- | --- |
| `sheet_name` | `string` | Cả hai | Tên sheet, mặc định là `Sheet1` |
| `overwrite` | `boolean` | Cả hai | Ghi đè nếu file đã tồn tại |
| `template_path` | `string` | Cả hai | Template `.xlsx` để mở trước khi ghi dữ liệu |
| `workbook_json` | `string` | Nâng cao | JSON string mô tả workbook |
| `headers_json` | `string` | Đơn giản | JSON string mảng tiêu đề cột |
| `rows_json` | `string` | Đơn giản | JSON string mảng các dòng dữ liệu |
| `column_widths_json` | `string` | Đơn giản | JSON string object cấu hình độ rộng cột |
| `freeze_header` | `boolean` | Đơn giản | Freeze hàng tiêu đề |
| `auto_filter` | `boolean` | Đơn giản | Bật autofilter cho header |
| `header_fill_color` | `string` | Đơn giản | Màu nền header |
| `header_font_color` | `string` | Đơn giản | Màu chữ header |
| `header_bold` | `boolean` | Đơn giản | In đậm header |
| `header_border` | `boolean` | Đơn giản | Viền mỏng cho header |
| `body_wrap_text` | `boolean` | Đơn giản | Wrap text cho body |
| `body_border` | `boolean` | Đơn giản | Viền mỏng cho body |
| `body_number_format` | `string` | Đơn giản | Excel custom number format cho body, ví dụ `#,##0` hoặc `#,##0.00` |

### Output

| Field | Kiểu | Ý nghĩa |
| --- | --- | --- |
| `output_path` | `string` | Đường dẫn tuyệt đối tới file đã lưu |
| `sheet_name` | `string` | Tên sheet đã ghi |
| `bytes_written` | `integer` | Kích thước file |
| `used_template` | `boolean` | Có dùng template hay không |
| `task` | `string` | Giá trị task đã gửi khi upload |
| `upload_message` | `string` | Message từ API upload |
| `upload_status` | `integer` | Status trong response upload |
| `download_url` | `string` | URL tải file sau khi upload |

Tool trả structured output theo MCP `structuredContent`. Với các client như Dify, các field trên có thể được expose thành output variables riêng; phần `text/files/json` là lớp hiển thị của client, không phải toàn bộ MCP result raw.

## Hai cách sử dụng chính

### 1. Chế độ đơn giản

Dùng khi chỉ cần tạo sheet dạng bảng với header và rows.

Bạn truyền:

- `headers_json`
- `rows_json`
- tùy chọn thêm style, freeze header, auto filter, column widths

Các trường này là `string`, nên giá trị thực tế phải là JSON đã stringify.

Ví dụ đối số:

```json
{
  "output_path": "demo/price-report.xlsx",
  "task": "price_report",
  "sheet_name": "Prices",
  "headers_json": "[\"Product\",\"Unit Price\",\"Quantity\"]",
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

#### `rows_json` hỗ trợ 2 dạng

##### Dạng mảng các mảng

```json
[
  ["Apple", 12000, 5],
  ["Orange", 9000, 8]
]
```

##### Dạng mảng các object

```json
[
  {"Product": "Apple", "Unit Price": 12000, "Quantity": 5},
  {"Product": "Orange", "Unit Price": 9000, "Quantity": 8}
]
```

Hành vi khi dùng object rows:

- Nếu đã truyền `headers_json`, tool lấy đúng thứ tự header đó để map dữ liệu
- Nếu không truyền `headers_json`, tool tự suy ra header từ object đầu tiên
- Các key suy ra tự động sẽ được sort tăng dần theo alphabet

#### `column_widths_json`

Có thể dùng:

- Tên cột Excel: `"A"`, `"B"`, `"AA"`
- Hoặc đúng text của header

Ví dụ:

```json
{
  "A": 20,
  "Unit Price": 14
}
```

Lưu ý:

- Match theo header là so khớp chính xác chuỗi
- Nếu key không phải tên cột Excel hợp lệ và cũng không khớp header, tool sẽ báo lỗi

#### Style mặc định trong simple mode

Simple mode không cho truyền style map tùy ý. Tool sẽ tự tạo tối đa 2 style nội bộ:

- `header`
- `body`

`header` chỉ được tạo nếu có ít nhất một field trong nhóm:

- `header_fill_color`
- `header_font_color`
- `header_bold`
- `header_border`

`body` chỉ được tạo nếu có ít nhất một field trong nhóm:

- `body_wrap_text`
- `body_border`
- `body_number_format`

Ví dụ nếu muốn số hiển thị có phân tách hàng nghìn, có thể truyền:

```json
{
  "body_number_format": "#,##0"
}
```

Lưu ý:

- Excel hiển thị dấu phân tách theo locale hoặc system separators của máy mở file
- Vì vậy format code nên là `#,##0`; trên máy dùng locale Việt Nam, Excel thường sẽ hiển thị thành `3.000.000`

### 2. Chế độ nâng cao với `workbook_json`

Dùng khi cần kiểm soát trực tiếp cell, merge, row height, hyperlink, freeze panes và style theo từng cell.

`workbook_json` cũng là một JSON string. Bên trong string này là object có cấu trúc `WorkbookSpec`.

Ví dụ readable payload trước khi stringify:

```json
{
  "sheet_name": "Report",
  "column_widths": [
    {"from": "A", "to": "A", "width": 24},
    {"from": "B", "to": "C", "width": 16}
  ],
  "row_heights": [
    {"row": 1, "height": 28}
  ],
  "merges": [
    {"start": "A1", "end": "C1"}
  ],
  "freeze_panes": {
    "freeze": true,
    "y_split": 1,
    "top_left_cell": "A2",
    "active_pane": "bottomLeft"
  },
  "styles": {
    "title": {
      "fill_color": "#D9EAF7",
      "font_color": "#111827",
      "bold": true,
      "horizontal": "center",
      "vertical": "center",
      "border": true
    },
    "link": {
      "font_color": "#2563EB",
      "underline": true
    }
  },
  "cells": [
    {"ref": "A1", "value": "Price Report", "style_ref": "title"},
    {"ref": "A2", "value": "Open file", "style_ref": "link", "hyperlink": {"url": "https://example.com", "display": "Open file"}}
  ]
}
```

Ví dụ khi truyền cho tool:

```json
{
  "output_path": "demo/advanced-report.xlsx",
  "task": "advanced_report",
  "workbook_json": "{\"sheet_name\":\"Report\",\"column_widths\":[{\"from\":\"A\",\"to\":\"A\",\"width\":24},{\"from\":\"B\",\"to\":\"C\",\"width\":16}],\"row_heights\":[{\"row\":1,\"height\":28}],\"merges\":[{\"start\":\"A1\",\"end\":\"C1\"}],\"freeze_panes\":{\"freeze\":true,\"y_split\":1,\"top_left_cell\":\"A2\",\"active_pane\":\"bottomLeft\"},\"styles\":{\"title\":{\"fill_color\":\"#D9EAF7\",\"font_color\":\"#111827\",\"bold\":true,\"horizontal\":\"center\",\"vertical\":\"center\",\"border\":true},\"link\":{\"font_color\":\"#2563EB\",\"underline\":true}},\"cells\":[{\"ref\":\"A1\",\"value\":\"Price Report\",\"style_ref\":\"title\"},{\"ref\":\"A2\",\"value\":\"Open file\",\"style_ref\":\"link\",\"hyperlink\":{\"url\":\"https://example.com\",\"display\":\"Open file\"}}]}"
}
```

#### `workbook_json` hỗ trợ gì

| Thành phần | Ý nghĩa |
| --- | --- |
| `sheet_name` | Tên sheet |
| `column_widths` | Cấu hình độ rộng cột |
| `row_heights` | Cấu hình chiều cao dòng |
| `merges` | Merge cell ranges |
| `cells` | Danh sách cell cần ghi |
| `freeze_panes` | Cấu hình pane freezing |
| `styles` | Map tên style sang `StyleSpec` |

#### Cấu trúc `CellSpec`

| Field | Ý nghĩa |
| --- | --- |
| `ref` | Tọa độ ô, ví dụ `A1` |
| `value` | Giá trị ô. Nếu là string, có thể dùng `<b>...</b>` hoặc `<strong>...</strong>` để in đậm một phần text |
| `style_ref` | Tên style trong `styles` |
| `hyperlink` | Link ngoài cho cell |

#### Cấu trúc `StyleSpec`

| Field | Ý nghĩa |
| --- | --- |
| `fill_color` | Màu nền |
| `font_color` | Màu chữ |
| `bold` | In đậm |
| `underline` | Gạch chân |
| `font_size` | Cỡ chữ |
| `horizontal` | Canh ngang |
| `vertical` | Canh dọc |
| `wrap_text` | Wrap text |
| `number_format` | Excel custom number format, ví dụ `#,##0` hoặc `#,##0.00` |
| `border` | Bật viền 4 cạnh |
| `border_color` | Màu viền |

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

Ví dụ text có in đậm một phần:

```json
{
  "ref": "A1",
  "value": "Tổng tiền: <b>3.000.000</b> VND"
}
```

#### Quy tắc ưu tiên khi có `workbook_json`

Nếu `workbook_json` có giá trị, tool sẽ bỏ qua các trường build dữ liệu dạng đơn giản:

- `headers_json`
- `rows_json`
- `column_widths_json`
- `freeze_header`
- `auto_filter`
- toàn bộ nhóm style header/body của simple mode

`sheet_name` chỉ còn vai trò fallback nếu `workbook_json.sheet_name` bị bỏ trống.

## Template mode

Nếu truyền `template_path`, tool sẽ mở file template trước khi ghi dữ liệu.

Hành vi thực tế:

- `template_path` phải là file `.xlsx`
- Đường dẫn template bị ràng buộc trong `EXCEL_TEMPLATE_DIR`
- Nếu sheet đã tồn tại trong template, tool ghi đè các ô được chỉ định lên sheet đó
- Nếu sheet chưa tồn tại, tool sẽ tạo sheet mới
- Tool không có bước clear sheet cũ trước khi ghi
- Tool không xóa các sheet khác trong template

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

## Quy tắc đường dẫn và an toàn

Project dùng `resolvePathInsideBase` để chặn path traversal.

### `output_path`

- Bắt buộc
- Phải kết thúc bằng `.xlsx`
- Nếu là relative path thì được resolve bên trong `EXCEL_OUTPUT_DIR`
- Nếu là absolute path thì vẫn phải nằm trong `EXCEL_OUTPUT_DIR`
- Không được phép escape ra ngoài thư mục gốc

### `template_path`

- Không bắt buộc
- Phải kết thúc bằng `.xlsx`
- Nếu là relative path thì được resolve bên trong `EXCEL_TEMPLATE_DIR`
- Nếu là absolute path thì vẫn phải nằm trong `EXCEL_TEMPLATE_DIR`
- Không được phép escape ra ngoài thư mục gốc

Ví dụ sẽ bị từ chối:

```text
../../etc/passwd
../other/report.xlsx
```

## Luồng xử lý nội bộ

Đây là flow thực tế của `create_excel_file`:

```mermaid
flowchart TD
    A[Nhận input MCP] --> B[Validate task và output_path]
    B --> C[Tạo thư mục output nếu cần]
    C --> D[Kiểm tra overwrite]
    D --> E[Mở template hoặc tạo workbook mới]
    E --> F[Build dữ liệu sheet]
    F --> G[Đảm bảo sheet tồn tại]
    G --> H[Áp column widths, row heights, cells, merges, panes]
    H --> I[Áp auto filter nếu có]
    I --> J[Lưu file .xlsx]
    J --> K[Upload file qua multipart HTTP]
    K --> L[Trả output + download_url]
```

### Diễn giải từng bước

#### 1. Validate input cơ bản

Tool kiểm tra:

- `task` không được rỗng
- `output_path` hợp lệ và nằm trong thư mục được phép

Sau đó tool:

- tạo thư mục cha của file output nếu chưa có
- từ chối nếu file đã tồn tại và `overwrite=false`

#### 2. Mở workbook

- Nếu không có `template_path`, tool tạo workbook mới bằng `excelize.NewFile()`
- Nếu có `template_path`, tool mở template bằng `excelize.OpenFile()`

#### 3. Build sheet

Có hai nhánh:

- Nếu có `workbook_json`, parse trực tiếp thành `WorkbookSpec`
- Nếu không có `workbook_json`, parse `headers_json`, `rows_json`, `column_widths_json` rồi tự build `SheetSpec`

Trong simple mode:

- header được ghi ở dòng 1 nếu có
- body bắt đầu từ dòng 2 nếu có header, ngược lại bắt đầu từ dòng 1
- `freeze_header=true` sẽ tạo `FreezePanesSpec` với `y_split=1`
- `auto_filter=true` sẽ tạo vùng filter từ `A1` tới ô cuối cùng của bảng

#### 4. Đảm bảo sheet tồn tại

`ensureSheets()` có hành vi:

- Nếu workbook mới tạo, sheet mặc định đầu tiên sẽ được rename thành sheet yêu cầu
- Nếu workbook hoặc template chưa có sheet đó, tool tạo mới
- Nếu sheet đã tồn tại, tool giữ nguyên và tiếp tục ghi dữ liệu

#### 5. Ghi dữ liệu vào sheet

`applySheet()` thực hiện theo thứ tự:

1. Đặt độ rộng cột
2. Đặt chiều cao dòng
3. Ghi từng cell
4. Merge cell ranges
5. Cấu hình freeze panes

Mỗi cell được ghi qua `writeCell()`:

- gọi `SetCellValue`
- nếu có hyperlink thì gắn external hyperlink
- nếu có `style_ref` thì resolve style rồi áp dụng style cho cell

#### 6. Resolve style và cache style

`resolveStyleID()`:

- lấy `StyleSpec` từ map `styles`
- marshal style thành JSON để tạo cache key
- nếu style đã từng tạo thì tái sử dụng `styleID`
- nếu chưa có thì gọi `f.NewStyle(...)`

Điểm này giúp tránh tạo trùng nhiều style giống nhau trong cùng một workbook.

#### 7. Save và upload

Sau khi sheet đã được ghi:

- workbook được `SaveAs(outputPath)`
- tool `stat` file để lấy kích thước
- file được upload qua `multipart/form-data`

Form upload gửi 3 phần:

- `file`
- `date_in_url=false`
- `task=<giá trị task>`

Tool chỉ coi upload là thành công khi:

- HTTP status nằm trong khoảng `2xx`
- response parse được thành JSON
- `status == 1`
- `url` không rỗng

## Các file chính trong codebase

| File | Vai trò |
| --- | --- |
| `main.go` | Tạo MCP server, đăng ký tool, chọn HTTP hoặc stdio transport, health check, path safety helpers |
| `simple.go` | Tool handler `createExcelFile`, schema input/output, parse simple mode và advanced mode |
| `types.go` | Các struct input/output và spec |
| `sheet.go` | Mở workbook, tạo sheet, ghi cell, merge, freeze panes |
| `style.go` | Resolve style và cache `styleID` |
| `convert.go` | Chuyển `StyleSpec` và `FreezePanesSpec` sang kiểu của `excelize` |
| `upload.go` | Upload file sau khi tạo xong |

## Những hành vi cần biết để dùng đúng

- `workbook_json`, `headers_json`, `rows_json`, `column_widths_json` đều là string chứa JSON, không phải object/array native trong schema tool
- `workbook_json` phải có ít nhất một cell
- `rows_json` chỉ chấp nhận một trong hai dạng: mảng các mảng hoặc mảng các object
- Không được trộn 2 dạng row trong cùng một payload
- Nếu `auto_filter=true` nhưng không có header, tool sẽ không áp autofilter
- Nếu `freeze_header=true` nhưng không có header, tool sẽ không tạo freeze panes
- Nếu `style_ref` không tồn tại trong `styles`, tool sẽ báo lỗi
- Nếu upload thất bại, local file có thể đã tồn tại trên đĩa

## Lỗi thường gặp

| Tình huống | Hành vi |
| --- | --- |
| `task` rỗng | Báo lỗi ngay |
| `output_path` không phải `.xlsx` | Báo lỗi |
| Path thoát ra ngoài base dir | Báo lỗi |
| File đã tồn tại và `overwrite=false` | Báo lỗi |
| `workbook_json` parse lỗi | Báo lỗi |
| `workbook_json` không có `cells` | Báo lỗi |
| `rows_json` không đúng format | Báo lỗi |
| `column_widths_json` dùng key lạ | Báo lỗi |
| `template_path` không mở được | Báo lỗi |
| Upload API trả status không hợp lệ | Báo lỗi |

## Ví dụ output thành công

```json
{
  "output_path": "/absolute/path/to/out/demo/price-report.xlsx",
  "sheet_name": "Prices",
  "bytes_written": 6421,
  "used_template": false,
  "task": "price_report",
  "upload_message": "Upload successful",
  "upload_status": 1,
  "download_url": "https://example.com/files/price-report.xlsx"
}
```

## Tóm tắt nhanh

Nếu chỉ cần xuất bảng đơn giản, dùng:

- `headers_json`
- `rows_json`
- các cờ format cơ bản

Nếu cần kiểm soát chi tiết từng ô, dùng:

- `workbook_json`

Flow luôn là:

- build workbook
- lưu file local
- upload file
- trả `download_url`
