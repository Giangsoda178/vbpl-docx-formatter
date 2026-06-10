#!/usr/bin/env python3
"""
vbpl_reformat.py
================
Định dạng lại file Word (.docx) văn bản pháp luật Việt Nam đã có sẵn theo
chuẩn định dạng văn bản hành chính (Times New Roman 14pt, justify, thụt
dòng đầu 0.5", các tiêu đề Điều/Chương/Mục đậm, Căn cứ in nghiêng, v.v.).

Hỗ trợ các loại văn bản: Nghị định, Luật, Thông tư, Quyết định, Nghị quyết,
Chỉ thị.

CÁCH DÙNG
---------
    python vbpl_reformat.py <duong_dan_input.docx>
    python vbpl_reformat.py <duong_dan_input.docx> <duong_dan_output.docx>

Nếu không chỉ định output, file kết quả sẽ là <input>_formatted.docx
cùng thư mục với input.

VÍ DỤ
-----
    python vbpl_reformat.py "Nghị định 141.docx"
    → Xuất "Nghị định 141_formatted.docx"

YÊU CẦU
-------
    pip install python-docx

CƠ CHẾ
------
Script đọc file đầu vào, trích xuất:
  1. Metadata từ bảng quốc hiệu (cơ quan, số văn bản, ngày)
  2. Loại văn bản + tiêu đề
  3. Danh sách các đoạn "Căn cứ"
  4. Các Chương / Mục / Điều / khoản trong thân bài
  5. Khối Nơi nhận + chữ ký

Sau đó xây dựng lại văn bản mới với định dạng chuẩn.
"""

import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
except ImportError as exc:
    raise ImportError(
        "The python-docx package is required. Install it with: pip install python-docx"
    ) from exc


# ============================================================
# HẰNG SỐ ĐỊNH DẠNG — chỉnh ở đây để đổi style toàn văn bản
# ============================================================

DEFAULT_FONT = "Times New Roman"
SIZE_DEFAULT = 12       # Quốc hiệu, NGHỊ ĐỊNH/LUẬT, chữ ký
SIZE_BODY = 14          # Thân bài, Điều, Căn cứ, tiêu đề dài
SIZE_DOC_NUMBER = 13    # Số văn bản
SIZE_RECIPIENTS = 11    # Danh sách Nơi nhận

INDENT_FIRST = Cm(1.27)   # Thụt dòng đầu = 0.5"
SPACE_BEFORE = Pt(3)      # Khoảng cách trên mỗi đoạn body
SPACE_ZERO = Pt(0)

PAGE_WIDTH = Cm(21.59)    # 8.5"
PAGE_HEIGHT = Cm(27.94)   # 11"
MARGIN_TOP = Cm(2.22)
MARGIN_OTHER = Cm(2.54)

# Các loại văn bản nhận dạng được
LOAI_VAN_BAN_KEYWORDS = {
    'NGHỊ ĐỊNH', 'LUẬT', 'THÔNG TƯ', 'QUYẾT ĐỊNH',
    'CHỈ THỊ', 'NGHỊ QUYẾT', 'PHÁP LỆNH',
    'THÔNG TƯ LIÊN TỊCH',
}

# Các đoạn mở đầu (sẽ in nghiêng + thụt dòng)
CAN_CU_STARTERS = (
    'Căn cứ',
    'Theo đề nghị',
    'Xét đề nghị',
    'Trên cơ sở',
    'Chính phủ ban hành',
    'Quốc hội ban hành',
    'Bộ trưởng ban hành',
    'Thủ tướng Chính phủ ban hành',
)


# ============================================================
# HELPER — Xử lý XML mức thấp
# ============================================================

def _set_cell_borders_none(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'nil')
        tcBorders.append(b)
    _AFTER = {qn('w:shd'), qn('w:noWrap'), qn('w:tcMar'),
              qn('w:textDirection'), qn('w:tcFitText'),
              qn('w:vAlign'), qn('w:hideMark')}
    insert_at = len(tcPr)
    for i, child in enumerate(tcPr):
        if child.tag in _AFTER:
            insert_at = i
            break
    tcPr.insert(insert_at, tcBorders)


def _set_table_borders_none(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    tblBorders = OxmlElement('w:tblBorders')
    for name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        b = OxmlElement(f'w:{name}')
        b.set(qn('w:val'), 'nil')
        tblBorders.append(b)
    _AFTER = {qn('w:shd'), qn('w:tblLayout'), qn('w:tblCellMar'),
              qn('w:tblLook'), qn('w:tblCaption'), qn('w:tblDescription')}
    insert_at = len(tblPr)
    for i, child in enumerate(tblPr):
        if child.tag in _AFTER:
            insert_at = i
            break
    tblPr.insert(insert_at, tblBorders)
    for row in table.rows:
        for cell in row.cells:
            _set_cell_borders_none(cell)


def _add_page_number_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement('w:fldChar')
    begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.text = 'PAGE \\* MERGEFORMAT'
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _set_title_page(section):
    sectPr = section._sectPr
    if sectPr.find(qn('w:titlePg')) is None:
        titlePg = OxmlElement('w:titlePg')
        _AFTER = {qn('w:textDirection'), qn('w:bidi'), qn('w:rtlGutter'),
                  qn('w:docGrid'), qn('w:printerSettings'),
                  qn('w:sectPrChange')}
        insert_at = len(sectPr)
        for i, child in enumerate(sectPr):
            if child.tag in _AFTER:
                insert_at = i
                break
        sectPr.insert(insert_at, titlePg)


def _set_default_font(doc, font_name, size_pt):
    style = doc.styles['Normal']
    style.font.name = font_name
    style.font.size = Pt(size_pt)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rFonts.set(qn(attr), font_name)


def _add_run(paragraph, text, *, bold=False, italic=False, size=SIZE_BODY):
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = DEFAULT_FONT
    run.font.size = Pt(size)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), DEFAULT_FONT)
    return run


def _set_cant_split(table_row):
    """Không cho row bị tách qua 2 trang."""
    tr = table_row._tr
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        tr.insert(0, trPr)
    trPr.append(OxmlElement('w:cantSplit'))


# ============================================================
# TRÍCH XUẤT — Đọc dữ liệu có cấu trúc từ document gốc
# ============================================================

def _cell_lines(cell):
    """Lấy các dòng không trống trong cell. Tách cả theo paragraph và
    theo line break (<w:br/>) — vì nhiều file gốc dùng br thay vì
    paragraph riêng."""
    lines = []
    for p in cell.paragraphs:
        for line in p.text.split('\n'):
            line = line.strip()
            if line:
                lines.append(line)
    return lines


def extract_metadata(doc):
    """Trích cơ quan, số văn bản, địa điểm, ngày từ bảng quốc hiệu."""
    if not doc.tables:
        raise ValueError(
            "Không tìm thấy bảng quốc hiệu ở đầu văn bản. "
            "File có vẻ không phải văn bản pháp luật chuẩn.")

    table = doc.tables[0]
    if len(table.rows) < 2:
        raise ValueError("Bảng quốc hiệu cần có ít nhất 2 hàng.")

    # Hàng 1, ô trái: tên cơ quan (có thể kèm dòng "--------")
    co_quan_lines = _cell_lines(table.rows[0].cells[0])
    co_quan = next((l for l in co_quan_lines if not re.fullmatch(r'-+', l)), '')

    # Hàng 2, ô trái: "Số: 141/2026/NĐ-CP" hoặc "Luật số: 58/2024/QH15"
    so_text = ' '.join(_cell_lines(table.rows[1].cells[0]))
    m = re.match(r'^(.+?:)\s*(.+)$', so_text)
    if m:
        prefix_so = m.group(1) + ' '
        so_van_ban = m.group(2)
    else:
        prefix_so = 'Số: '
        so_van_ban = so_text

    # Hàng 2, ô phải: "Hà Nội, ngày 29 tháng 4 năm 2026"
    dia_ngay = ' '.join(_cell_lines(table.rows[1].cells[1]))
    m = re.match(r'^(.+?),\s*ngày\s+(.+)$', dia_ngay)
    if m:
        dia_diem = m.group(1).strip()
        ngay = m.group(2).strip()
    else:
        dia_diem = 'Hà Nội'
        ngay = dia_ngay.strip()

    return {
        'co_quan': co_quan,
        'so_van_ban': so_van_ban,
        'dia_diem': dia_diem,
        'ngay_ban_hanh': ngay,
        'prefix_so': prefix_so,
    }


def _is_para_italic(para):
    """Đoạn có ≥ nửa số run là italic không?"""
    runs = [r for r in para.runs if r.text.strip()]
    if not runs:
        return False
    italic_runs = [r for r in runs if r.italic]
    return len(italic_runs) >= len(runs) / 2


def _runs_with_format(para):
    """Lấy list (text, {'bold', 'italic'}) cho từng run của paragraph."""
    parts = []
    for run in para.runs:
        if run.text:
            parts.append((run.text, {
                'bold': bool(run.bold),
                'italic': bool(run.italic),
            }))
    return parts


def _has_mixed_inline_format(para):
    """Đoạn có pha trộn run italic và non-italic không?
    (vd: định nghĩa thuật ngữ '1. *Báo cáo...* là tài liệu...')."""
    runs = _runs_with_format(para)
    if len(runs) < 2:
        return False
    has_i = any(f['italic'] for _, f in runs)
    has_ni = any(not f['italic'] for _, f in runs)
    return has_i and has_ni


def _classify(text):
    """Phân loại 1 đoạn dựa trên nội dung text."""
    text = text.strip()
    if not text:
        return ('blank',)

    if text in LOAI_VAN_BAN_KEYWORDS:
        return ('loai_van_ban', text)

    m = re.fullmatch(r'Chương\s+([IVXLCDM]+)', text)
    if m:
        return ('chuong_num', m.group(1))

    m = re.match(r'^Mục\s+(\d+)\.\s+(.+)$', text)
    if m:
        return ('muc', m.group(1), m.group(2).strip())

    m = re.match(r'^Điều\s+(\d+)\.\s+(.+)$', text)
    if m:
        return ('dieu', m.group(1), m.group(2).strip())

    for starter in CAN_CU_STARTERS:
        if text.startswith(starter):
            return ('can_cu', text)

    return ('khoan', text)


def extract_body_items(doc):
    """Đi qua body của document, trả về list các mục có cấu trúc."""
    body_xml = doc.element.body

    # Liệt kê tất cả paragraphs + tables theo đúng thứ tự xuất hiện
    flow = []   # list of ('p', Paragraph) | ('t', None)
    for child in body_xml.iterchildren():
        if child.tag == qn('w:tbl'):
            flow.append(('t', None))
        elif child.tag == qn('w:p'):
            flow.append(('p', Paragraph(child, doc.part)))

    # Xác định ranh giới: bỏ qua bảng đầu (quốc hiệu) và bảng cuối (chữ ký)
    table_idxs = [i for i, (k, _) in enumerate(flow) if k == 't']
    if not table_idxs:
        raise ValueError("Không tìm thấy bảng quốc hiệu.")

    start = table_idxs[0] + 1
    end = table_idxs[-1] if len(table_idxs) > 1 else len(flow)

    body_paras = [item[1] for item in flow[start:end] if item[0] == 'p']

    # Phân loại với context (chương cần chờ tên ở đoạn kế tiếp)
    items = []
    pending_chuong_num = None
    pending_loai = None

    for para in body_paras:
        text = para.text.strip()
        if not text:
            continue

        # Đang chờ tên chương?
        if pending_chuong_num is not None:
            items.append(('chuong', pending_chuong_num, text))
            pending_chuong_num = None
            continue

        # Đang chờ tiêu đề (sau loại văn bản)?
        if pending_loai is not None:
            items.append(('title', pending_loai, text))
            pending_loai = None
            continue

        kind = _classify(text)
        tag = kind[0]

        if tag == 'loai_van_ban':
            pending_loai = kind[1]
        elif tag == 'chuong_num':
            pending_chuong_num = kind[1]
        elif tag in ('muc', 'dieu', 'can_cu'):
            items.append(kind)
        else:  # 'khoan'
            # Kiểm tra inline italic / mixed formatting
            if _has_mixed_inline_format(para):
                items.append(('mixed', _runs_with_format(para)))
            elif _is_para_italic(para):
                items.append(('italic_body', text))
            else:
                items.append(kind)

    # Nếu còn pending — xử lý fallback
    if pending_loai is not None:
        items.append(('title', pending_loai, ''))
    if pending_chuong_num is not None:
        items.append(('chuong', pending_chuong_num, ''))

    return items


def extract_closing(doc):
    """Trích Nơi nhận, chức danh, người ký từ bảng cuối."""
    if len(doc.tables) < 2:
        return None

    table = doc.tables[-1]
    if not table.rows or len(table.rows[0].cells) < 2:
        return None

    # Ô trái: danh sách nơi nhận
    left = table.rows[0].cells[0]
    left_text = '\n'.join(p.text for p in left.paragraphs)

    # Chỉ tách theo newline — KHÔNG tách trên ' - ' inline
    # (vì dấu ' - ' có thể xuất hiện trong tên nơi nhận, vd:
    # 'tổ chức chính trị - xã hội')
    recipients = []
    for line in left_text.split('\n'):
        line = line.strip()
        if not line or 'Nơi nhận' in line:
            continue
        # Bỏ dấu '-' đầu (cũng chấp nhận trường hợp thiếu khoảng trắng: "-Văn phòng...")
        line = re.sub(r'^[-–—]\s*', '', line).strip()
        # Bỏ dấu ';' hoặc '.' cuối
        line = re.sub(r'[;.]\s*$', '', line).strip()
        # Bỏ ký tự  no-break space
        line = line.replace('\xa0', ' ')
        # Gộp nhiều khoảng trắng
        line = re.sub(r'\s{2,}', ' ', line)
        if line:
            recipients.append(line)

    # Ô phải: chức danh (in hoa) + chữ ký + tên người ký
    right = table.rows[0].cells[1]
    # Tách theo cả paragraph và line break
    right_lines = []
    for p in right.paragraphs:
        for line in p.text.split('\n'):
            line = line.strip()
            if line:
                right_lines.append(line)

    position_lines = []
    signer = ''
    for line in right_lines:
        # In hoa toàn bộ ⇒ chức danh
        if line.upper() == line and len(line) > 1:
            position_lines.append(line)
        else:
            # Dòng có chữ thường ⇒ tên người ký (lấy dòng đầu tiên)
            if not signer:
                signer = line

    return {
        'recipients': recipients,
        'position_lines': position_lines,
        'signer': signer,
    }


# ============================================================
# DỰNG LẠI VĂN BẢN — Tạo document mới với định dạng chuẩn
# ============================================================

def _add_header_table(doc, co_quan, so_van_ban, dia_diem,
                      ngay_ban_hanh, prefix_so):
    table = doc.add_table(rows=2, cols=2)
    _set_table_borders_none(table)
    table.columns[0].width = Cm(6.6)
    table.columns[1].width = Cm(9.9)

    # Hàng 1 trái: tên cơ quan
    left = table.cell(0, 0)
    left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(left.paragraphs[0], co_quan, bold=True, size=SIZE_DEFAULT)
    p = left.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(p, '--------', bold=True, size=SIZE_DEFAULT)

    # Hàng 1 phải: quốc hiệu
    right = table.cell(0, 1)
    right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(right.paragraphs[0],
             'CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM',
             bold=True, size=SIZE_DEFAULT)
    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(p, 'Độc lập - Tự do - Hạnh phúc', bold=True, size=SIZE_DEFAULT)
    p = right.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(p, '---------------', bold=True, size=SIZE_DEFAULT)

    # Hàng 2 trái: số văn bản
    left2 = table.cell(1, 0)
    left2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_run(left2.paragraphs[0], prefix_so + so_van_ban, size=SIZE_DOC_NUMBER)

    # Hàng 2 phải: địa điểm, ngày
    right2 = table.cell(1, 1)
    right2.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _add_run(right2.paragraphs[0],
             f'{dia_diem}, ngày {ngay_ban_hanh}',
             italic=True, size=SIZE_BODY)


def _add_title_block(doc, loai_van_ban, tieu_de):
    # Đoạn trống
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    _add_run(p, ' ', bold=True, size=SIZE_DEFAULT)

    # Loại văn bản
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    _add_run(p, loai_van_ban, bold=True, size=SIZE_DEFAULT)

    # Tiêu đề dài
    if tieu_de:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = SPACE_BEFORE
        _add_run(p, tieu_de, bold=True, size=SIZE_BODY)


def _add_can_cu(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, text, italic=True, size=SIZE_BODY)


def _add_chuong(doc, roman, name):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, f'Chương {roman}', bold=True, size=SIZE_BODY)

    if name:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = SPACE_BEFORE
        _add_run(p, name.upper(), bold=True, size=SIZE_BODY)


def _add_muc(doc, so, name):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, f'Mục {so}. {name.upper()}', bold=True, size=SIZE_BODY)


def _add_dieu(doc, so, name):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, f'Điều {so}. {name}', bold=True, size=SIZE_BODY)


def _add_khoan(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, text, size=SIZE_BODY)


def _add_italic_body(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    _add_run(p, text, italic=True, size=SIZE_BODY)


def _add_mixed(doc, parts):
    """Đoạn body với nhiều run định dạng khác nhau (giữ inline italic/bold)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = INDENT_FIRST
    p.paragraph_format.space_before = SPACE_BEFORE
    for text, fmt in parts:
        _add_run(p, text,
                 bold=fmt.get('bold', False),
                 italic=fmt.get('italic', False),
                 size=SIZE_BODY)


def _add_closing_table(doc, recipients, position_lines, signer,
                       blank_lines=5):
    # Đoạn trống ngăn cách
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    _add_run(p, ' ', size=SIZE_DEFAULT)

    table = doc.add_table(rows=1, cols=2)
    _set_table_borders_none(table)
    _set_cant_split(table.rows[0])
    table.columns[0].width = Cm(8.25)
    table.columns[1].width = Cm(8.25)

    # Ô trái: Nơi nhận
    left = table.cell(0, 0)
    p = left.paragraphs[0]
    p.paragraph_format.space_before = SPACE_ZERO
    p.paragraph_format.space_after = SPACE_ZERO
    if recipients:
        _add_run(p, 'Nơi nhận:', bold=True, italic=True, size=SIZE_RECIPIENTS)
        for line in recipients:
            p = left.add_paragraph()
            p.paragraph_format.space_before = SPACE_ZERO
            p.paragraph_format.space_after = SPACE_ZERO
            text = line.strip()
            if not text.startswith('-'):
                text = '- ' + text
            if not text.endswith(';') and not text.endswith('.'):
                text = text + ';'
            _add_run(p, text, size=SIZE_RECIPIENTS)

    # Ô phải: chức danh + chữ ký
    right = table.cell(0, 1)
    first = True
    for line in position_lines:
        p = right.paragraphs[0] if first else right.add_paragraph()
        first = False
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = SPACE_ZERO
        p.paragraph_format.space_after = SPACE_ZERO
        _add_run(p, line, bold=True, size=SIZE_BODY)
    # Dòng trống chừa cho chữ ký
    for _ in range(blank_lines):
        p = right.add_paragraph() if not first else right.paragraphs[0]
        first = False
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = SPACE_ZERO
        p.paragraph_format.space_after = SPACE_ZERO
        _add_run(p, ' ', size=SIZE_BODY)
    # Tên người ký
    if signer:
        p = right.add_paragraph() if not first else right.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = SPACE_ZERO
        p.paragraph_format.space_after = SPACE_ZERO
        _add_run(p, signer, bold=True, size=SIZE_BODY)


def build_document(metadata, body_items, closing):
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.top_margin = MARGIN_TOP
    section.bottom_margin = MARGIN_OTHER
    section.left_margin = MARGIN_OTHER
    section.right_margin = MARGIN_OTHER
    section.header_distance = Cm(1.27)
    section.footer_distance = Cm(1.27)
    _set_title_page(section)

    # Page number ở header
    p = section.header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_number_field(p)

    _set_default_font(doc, DEFAULT_FONT, SIZE_DEFAULT)

    # Quốc hiệu
    _add_header_table(doc, **metadata)

    # Body
    dispatch = {
        'title':       lambda it: _add_title_block(doc, it[1], it[2]),
        'can_cu':      lambda it: _add_can_cu(doc, it[1]),
        'chuong':      lambda it: _add_chuong(doc, it[1], it[2]),
        'muc':         lambda it: _add_muc(doc, it[1], it[2]),
        'dieu':        lambda it: _add_dieu(doc, it[1], it[2]),
        'khoan':       lambda it: _add_khoan(doc, it[1]),
        'italic_body': lambda it: _add_italic_body(doc, it[1]),
        'mixed':       lambda it: _add_mixed(doc, it[1]),
    }
    for item in body_items:
        handler = dispatch.get(item[0])
        if handler:
            handler(item)

    # Closing
    if closing and (closing['recipients'] or closing['position_lines']
                    or closing['signer']):
        _add_closing_table(doc, **closing)

    # Vá lỗi zoom của python-docx
    zoom = doc.settings.element.find(qn('w:zoom'))
    if zoom is not None and zoom.get(qn('w:percent')) is None:
        zoom.set(qn('w:percent'), '100')

    return doc


# ============================================================
# MAIN
# ============================================================

def reformat(input_path, output_path, verbose=True):
    """Đọc input docx → trích xuất → dựng lại theo chuẩn → lưu output."""
    if verbose:
        print(f"📖 Đọc: {input_path}")
    src = Document(str(input_path))

    metadata = extract_metadata(src)
    body_items = extract_body_items(src)
    closing = extract_closing(src)

    if verbose:
        print(f"   ▸ Cơ quan       : {metadata['co_quan']}")
        print(f"   ▸ Số văn bản    : {metadata['prefix_so']}{metadata['so_van_ban']}")
        print(f"   ▸ Ngày ban hành : {metadata['ngay_ban_hanh']}")
        # Đếm các loại item
        counts = {}
        for it in body_items:
            counts[it[0]] = counts.get(it[0], 0) + 1
        print(f"   ▸ Nội dung phát hiện:")
        labels = {
            'title': 'Tiêu đề', 'can_cu': 'Căn cứ', 'chuong': 'Chương',
            'muc': 'Mục', 'dieu': 'Điều', 'khoan': 'Khoản/đoạn',
            'italic_body': 'Đoạn in nghiêng', 'mixed': 'Đoạn pha định dạng',
        }
        for tag, count in counts.items():
            print(f"       - {labels.get(tag, tag):20s}: {count}")
        if closing:
            print(f"   ▸ Khối kết      : {len(closing['recipients'])} nơi nhận, "
                  f"người ký: {closing['signer'] or '(không)'}")
        else:
            print(f"   ▸ Khối kết      : không có")

    doc = build_document(metadata, body_items, closing)
    doc.save(str(output_path))
    if verbose:
        print(f"✅ Đã xuất: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description='Định dạng lại văn bản pháp luật Việt Nam (.docx)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ví dụ:\n"
            "  python vbpl_reformat.py 'Nghị định 141.docx'\n"
            "  python vbpl_reformat.py input.docx output.docx\n"
        ),
    )
    parser.add_argument('input', help='Đường dẫn file .docx cần định dạng')
    parser.add_argument('output', nargs='?',
                        help='File xuất (mặc định: <input>_formatted.docx)')
    parser.add_argument('-q', '--quiet', action='store_true',
                        help='Không in thông tin chi tiết')
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    if not input_path.exists():
        print(f"❌ Không tìm thấy file: {input_path}", file=sys.stderr)
        sys.exit(1)
    if input_path.suffix.lower() != '.docx':
        print(f"⚠️  Cảnh báo: File không có đuôi .docx",
              file=sys.stderr)

    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = (input_path.parent /
                       f"{input_path.stem}_formatted.docx")

    try:
        reformat(input_path, output_path, verbose=not args.quiet)
    except ValueError as e:
        print(f"❌ Lỗi khi xử lý: {e}", file=sys.stderr)
        sys.exit(2)
    except Exception as e:
        print(f"❌ Lỗi không xác định: {e}", file=sys.stderr)
        raise


if __name__ == '__main__':
    main()
