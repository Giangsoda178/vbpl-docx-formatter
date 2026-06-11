# vbpl_reformat.py

Định dạng lại file Word (.docx) văn bản pháp luật Việt Nam theo chuẩn (Times New Roman 14pt, justify, thụt dòng 0.5", v.v.).

## Yêu cầu

- Python 3.9+
- Gói `python-docx`

Trên macOS với Python cài qua Homebrew, hệ thống bị khóa cài gói toàn cục (PEP 668), nên dùng **venv** là cách an toàn nhất.

## Cài đặt (lần đầu)

```bash
# Tạo venv cố định trong thư mục home
python3 -m venv ~/.venvs/vbpl

# Cài python-docx vào venv
~/.venvs/vbpl/bin/pip install python-docx
```

## Chạy script

```bash
# Mặc định: xuất ra <input>_formatted.docx cùng thư mục
~/.venvs/vbpl/bin/python ~/Downloads/vbpl_reformat.py "<đường_dẫn_input.docx>"

# Chỉ định file output
~/.venvs/vbpl/bin/python ~/Downloads/vbpl_reformat.py "<input.docx>" "<output.docx>"

# Tắt log chi tiết
~/.venvs/vbpl/bin/python ~/Downloads/vbpl_reformat.py -q "<input.docx>"
```

### Ví dụ

```bash
~/.venvs/vbpl/bin/python ~/Downloads/vbpl_reformat.py \
    "~/Desktop/VBPL/7_2026_TT-BNV_704494.docx"
```

Kết quả: `~/Desktop/VBPL/7_2026_TT-BNV_704494_formatted.docx`

## Rút gọn lệnh (tùy chọn)

Thêm alias vào `~/.zshrc`:

```bash
alias vbpl='~/.venvs/vbpl/bin/python ~/Downloads/vbpl_reformat.py'
```

Sau khi `source ~/.zshrc`:

```bash
vbpl "~/Desktop/VBPL/7_2026_TT-BNV_704494.docx"
```

## Loại văn bản hỗ trợ

Nghị định, Luật, Thông tư, Quyết định, Chỉ thị, Nghị quyết, Pháp lệnh, Thông tư liên tịch.

## Xử lý sự cố

- **`ModuleNotFoundError: No module named 'docx'`** — venv chưa được kích hoạt hoặc chưa cài `python-docx`. Chạy lại bước cài đặt.
- **`error: externally-managed-environment`** khi cài bằng `pip` trực tiếp — đúng như mong đợi trên macOS Homebrew. Dùng venv như hướng dẫn ở trên thay vì `pip install` toàn cục.
- **`Không tìm thấy bảng quốc hiệu`** — file input không có bảng quốc hiệu ở đầu (không đúng định dạng văn bản pháp luật chuẩn).
