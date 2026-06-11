# VBPL DOCX Formatter

Web app for reformatting Vietnamese legal documents (Nghị định, Luật, Thông tư, Quyết định, Nghị quyết, Chỉ thị) in `.docx` format to the standard administrative document style (Times New Roman 14pt, justified, first-line indent, bold Điều/Chương/Mục headings, italic Căn cứ, …).

Upload a `.docx` in the browser → the backend runs [vbpl_reformat.py](vbpl_reformat.py) → download the formatted result.

## Architecture

- **[client/](client/)** — Vite + TypeScript single-page frontend (drag-and-drop upload, download)
- **[server/](server/)** — Express + TypeScript API; `POST /api/format` writes the upload to a temp dir, spawns the Python script, streams back the formatted file, and cleans up
- **[vbpl_reformat.py](vbpl_reformat.py)** — the formatting engine (python-docx)

## Prerequisites

- Node.js ≥ 20
- Python 3 with `python-docx` (`pip install python-docx`)

## Setup

```sh
npm install
```

## Development

```sh
npm run dev
```

Runs the API server on http://localhost:3001 and the Vite dev server (with `/api` proxy) on http://localhost:5173. Open the latter.

## Production

```sh
npm run build
npm run start
```

Builds the client and server, then serves both from http://localhost:3001.

Environment variables:

| Variable     | Default                                  | Purpose                    |
| ------------ | ---------------------------------------- | -------------------------- |
| `PORT`       | `3001`                                   | Server port                |
| `PYTHON_BIN` | `python` (Windows) / `python3` (others)  | Python interpreter to use  |

## Testing

With the server running:

```sh
python test/make_sample.py      # generate a sample legal document
node test/test_api.mjs          # end-to-end API tests
```

## API

`POST /api/format` — multipart form with a `file` field containing a `.docx`.

- `200` — formatted document (`Content-Disposition` carries `<name>_formatted.docx`)
- `400` — missing file, wrong type, or file over 20 MB
- `422` — document structure not recognized (e.g. missing the national-emblem header table) or corrupt file; body is `{ "error": "<message>" }`
- `500` — unexpected processing error
