# Connections

Registry của mọi hệ thống AIOS có thể kết nối. Cập nhật mỗi khi kết nối thêm công cụ mới.

| # | Domain | Công cụ | Cơ chế | Auth | Kiểm tra lần cuối |
|---|---|---|---|---|---|
| 1 | Tài chính / Thu nhập | Tài khoản ngân hàng | not yet connected | — | — |
| 2 | Hồ sơ / Tài liệu | Dropbox | not yet connected | — | — |
| 3 | Hồ sơ / Tài liệu | Google Drive | not yet connected | — | — |
| 4 | Lịch công tác | Google Calendar (qua Gmail) | not yet connected | — | — |
| 5 | Liên lạc | Zalo | not yet connected | — | — |
| 6 | Liên lạc | Điện thoại | manual | — | — |
| 7 | Quản lý công việc | (chưa có công cụ riêng) | not yet connected | — | — |

**Cơ chế:** `mcp` (MCP server), `script` (Python/Bash gọi API, trong `scripts/`), `export` (dump CSV/JSON), `key+ref` (API key + `references/{tool}-api.md`), `not yet connected`, `manual`.

Khi kết nối thêm công cụ, lưu thêm `references/{tool}-api.md` ghi lại endpoints, cách xác thực và các truy vấn thường dùng.
