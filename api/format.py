"""Vercel serverless function: POST /api/format

Same contract as the Express route in server/src/index.ts (multipart upload,
docx response, 400/422/500 JSON errors), but runs vbpl_reformat in-process —
no child Python process. Deployed automatically by Vercel from api/format.py;
run locally via test/serve_api_local.py.
"""

import json
import os
import sys
import tempfile
import traceback
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import quote

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from docx.opc.exceptions import PackageNotFoundError

from vbpl_reformat import reformat

DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
# Vercel rejects bigger request bodies with 413 before the function runs;
# this cap only matters when the handler is served locally.
MAX_FILE_SIZE = int(4.5 * 1024 * 1024)

MSG_NO_FILE = 'Chưa chọn file. Vui lòng tải lên một file .docx.'
MSG_BAD_TYPE = 'File không hợp lệ. Chỉ chấp nhận file Word (.docx).'
MSG_TOO_LARGE = 'File quá lớn. Kích thước tối đa là 4,5 MB.'
MSG_CORRUPT = 'File bị hỏng hoặc không phải định dạng Word (.docx) hợp lệ.'
MSG_SERVER_ERROR = 'Lỗi máy chủ khi xử lý file.'


def _fix_header_utf8(value):
    """Browsers put raw UTF-8 bytes in multipart Content-Disposition filenames;
    the email parser surfaces those bytes as surrogate escapes. Recover the
    original UTF-8 string (mirrors the latin1 fix-up in server/src/index.ts)."""
    if not value:
        return value
    try:
        return value.encode('ascii', 'surrogateescape').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return value


def _extract_upload(content_type, body):
    """Return (filename, mimetype, data) for the 'file' field, or None."""
    if not content_type or 'multipart/form-data' not in content_type.lower():
        return None
    msg = BytesParser(policy=policy.default).parsebytes(
        b'Content-Type: ' + content_type.encode('latin-1') + b'\r\n\r\n' + body
    )
    if not msg.is_multipart():
        return None
    for part in msg.iter_parts():
        if part.get_param('name', header='content-disposition') != 'file':
            continue
        filename = _fix_header_utf8(part.get_filename() or '')
        data = part.get_payload(decode=True) or b''
        return filename, part.get_content_type(), data
    return None


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            length = int(self.headers.get('content-length') or 0)
            if length > MAX_FILE_SIZE + 4096:  # headroom for multipart framing
                self._send_json(400, MSG_TOO_LARGE)
                return
            body = self.rfile.read(length)

            upload = _extract_upload(self.headers.get('content-type'), body)
            if upload is None:
                self._send_json(400, MSG_NO_FILE)
                return
            filename, mimetype, data = upload
            if not filename and not data:
                self._send_json(400, MSG_NO_FILE)
                return
            if not (filename.lower().endswith('.docx') or mimetype == DOCX_MIME):
                self._send_json(400, MSG_BAD_TYPE)
                return
            if len(data) > MAX_FILE_SIZE:
                self._send_json(400, MSG_TOO_LARGE)
                return

            with tempfile.TemporaryDirectory(prefix='vbpl-') as work_dir:
                input_path = os.path.join(work_dir, 'input.docx')
                output_path = os.path.join(work_dir, 'output.docx')
                with open(input_path, 'wb') as f:
                    f.write(data)
                try:
                    reformat(input_path, output_path, verbose=False)
                except PackageNotFoundError:
                    self._send_json(422, MSG_CORRUPT)
                    return
                except ValueError as exc:
                    self._send_json(422, f'Lỗi khi xử lý: {exc}')
                    return
                with open(output_path, 'rb') as f:
                    result = f.read()

            stem = Path(filename or 'document').stem
            self._send_docx(result, f'{stem}_formatted.docx')
        except Exception:
            traceback.print_exc()
            self._send_json(500, MSG_SERVER_ERROR)

    def _send_json(self, status, message):
        payload = json.dumps({'error': message}, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _send_docx(self, data, download_name):
        # RFC 5987/6266: ASCII-safe fallback plus UTF-8 encoded real name.
        encoded = quote(download_name, safe='')
        ascii_name = download_name.encode('ascii', 'ignore').decode('ascii') or 'formatted.docx'
        self.send_response(200)
        self.send_header('Content-Type', DOCX_MIME)
        self.send_header(
            'Content-Disposition',
            f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{encoded}',
        )
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)
