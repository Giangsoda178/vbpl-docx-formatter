# VBPL DOCX Formatter

Web app for reformatting Vietnamese legal documents (Nghị định, Luật, Thông tư, Quyết định, Nghị quyết, Chỉ thị) in `.docx` format to the standard administrative document style (Times New Roman 14pt, justified, first-line indent, bold Điều/Chương/Mục headings, italic Căn cứ, …).

Upload a `.docx` in the browser → the backend runs [vbpl_reformat.py](vbpl_reformat.py) → download the formatted result.

## Architecture

- **[client/](client/)** — Vite + TypeScript single-page frontend (drag-and-drop upload, download)
- **[server/](server/)** — Express + TypeScript API; `POST /api/format` writes the upload to a temp dir, spawns the Python script, streams back the formatted file, and cleans up (used for local dev / self-hosting)
- **[api/format.py](api/format.py)** — the same `POST /api/format` endpoint as a Vercel Python serverless function; calls the formatting engine in-process (used on Vercel)
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

## Deploying to Vercel

The repo is Vercel-ready. The Express server is not used on Vercel; instead
[api/format.py](api/format.py) runs as a Python serverless function and the built client is served
statically, per [vercel.json](vercel.json):

- `buildCommand: vite build client`, `outputDirectory: client/dist`, with an SPA fallback rewrite
  (`/api/*` is unaffected — functions take filesystem precedence over rewrites)
- `api/format.py` is auto-deployed at `/api/format`; Python dependencies come from
  [requirements.txt](requirements.txt)
- `maxDuration: 60` and `includeFiles: vbpl_reformat.py` are set for the function

To deploy, either:

1. Push the repo to GitHub/GitLab/Bitbucket and import it at [vercel.com/new](https://vercel.com/new)
   (keep the Root Directory at the repo root — `api/` must be visible to Vercel), or
2. Use the CLI: `npm i -g vercel`, then `vercel` from the repo root (`vercel --prod` to go live).

**Limits on Vercel:** request and response bodies are capped at 4.5 MB (uploads above that get
HTTP 413 before reaching the function — the client shows a friendly message). This cap cannot be
raised on any plan; if larger files are ever needed, the upload would have to go through
[Vercel Blob](https://vercel.com/docs/vercel-blob). Typical legal documents are far below the cap.

To test the exact Vercel handler locally (no deploy needed):

```sh
python test/serve_api_local.py                          # serves api/format.py on :3002
API_BASE=http://127.0.0.1:3002 node test/test_api.mjs   # e2e suite against it
python test/verify_vercel_handler.py                    # deep checks (docx validity, 422 path)
```

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
