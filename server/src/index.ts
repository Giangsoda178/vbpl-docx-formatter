import { spawn } from 'node:child_process';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import express from 'express';
import multer from 'multer';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = path.resolve(__dirname, '..', '..');
const SCRIPT_PATH = path.join(PROJECT_ROOT, 'vbpl_reformat.py');
const PYTHON_BIN = process.env.PYTHON_BIN ?? (process.platform === 'win32' ? 'python' : 'python3');
const PORT = Number(process.env.PORT ?? 3001);
const MAX_FILE_SIZE = 4.5 * 1024 * 1024; // 4.5 MB (Vercel request body limit)

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_FILE_SIZE },
});

interface ScriptResult {
  code: number | null;
  stderr: string;
}

function runReformat(inputPath: string, outputPath: string): Promise<ScriptResult> {
  return new Promise((resolve, reject) => {
    const proc = spawn(PYTHON_BIN, [SCRIPT_PATH, inputPath, outputPath, '--quiet'], {
      cwd: PROJECT_ROOT,
      env: { ...process.env, PYTHONIOENCODING: 'utf-8' },
      windowsHide: true,
    });
    let stderr = '';
    proc.stderr.on('data', (chunk: Buffer) => {
      stderr += chunk.toString('utf-8');
    });
    proc.on('error', reject);
    proc.on('close', (code) => resolve({ code, stderr }));
  });
}

// Multer decodes originalname as latin1; recover the UTF-8 name (Vietnamese filenames).
function decodeOriginalName(name: string): string {
  return Buffer.from(name, 'latin1').toString('utf-8');
}

// Strip the emoji/log prefixes the script prints so the client gets a clean
// message, and drop any Python traceback (internal detail, not user-facing).
function cleanErrorMessage(stderr: string): string {
  const tracebackStart = stderr.indexOf('Traceback (most recent call last)');
  const visible = tracebackStart >= 0 ? stderr.slice(0, tracebackStart) : stderr;
  return visible
    .split(/\r?\n/)
    .map((line) => line.replace(/^[❌⚠️\s]+/u, '').trim())
    .filter(Boolean)
    .join('\n');
}

const app = express();

app.post('/api/format', upload.single('file'), async (req, res) => {
  const file = req.file;
  if (!file) {
    res.status(400).json({ error: 'Chưa chọn file. Vui lòng tải lên một file .docx.' });
    return;
  }

  const originalName = decodeOriginalName(file.originalname);
  const looksLikeDocx = originalName.toLowerCase().endsWith('.docx') || file.mimetype === DOCX_MIME;
  if (!looksLikeDocx) {
    res.status(400).json({ error: 'File không hợp lệ. Chỉ chấp nhận file Word (.docx).' });
    return;
  }

  let workDir: string | null = null;
  try {
    workDir = await mkdtemp(path.join(tmpdir(), 'vbpl-'));
    const inputPath = path.join(workDir, 'input.docx');
    const outputPath = path.join(workDir, 'output.docx');
    await writeFile(inputPath, file.buffer);

    const { code, stderr } = await runReformat(inputPath, outputPath);

    if (code !== 0 || !existsSync(outputPath)) {
      const message = cleanErrorMessage(stderr) || 'Xử lý file thất bại.';
      // Exit code 2 = the script rejected the document structure (user-fixable).
      res.status(code === 2 ? 422 : 500).json({ error: message });
      return;
    }

    const stem = path.basename(originalName, path.extname(originalName));
    const downloadName = `${stem}_formatted.docx`;
    res.download(outputPath, downloadName, (err) => {
      void cleanup(workDir);
      if (err && !res.headersSent) {
        res.status(500).json({ error: 'Không gửi được file kết quả.' });
      }
    });
    workDir = null; // ownership transferred to the download callback
  } catch (err) {
    console.error('Format request failed:', err);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Lỗi máy chủ khi xử lý file.' });
    }
  } finally {
    void cleanup(workDir);
  }
});

async function cleanup(dir: string | null) {
  if (!dir) return;
  try {
    await rm(dir, { recursive: true, force: true });
  } catch {
    // best effort — temp dir will be reclaimed by the OS eventually
  }
}

// Multer errors (e.g. file too large) and other middleware errors.
app.use((err: unknown, _req: express.Request, res: express.Response, _next: express.NextFunction) => {
  if (err instanceof multer.MulterError) {
    const message =
      err.code === 'LIMIT_FILE_SIZE'
        ? `File quá lớn. Kích thước tối đa là ${MAX_FILE_SIZE / 1024 / 1024} MB.`
        : `Lỗi tải file: ${err.message}`;
    res.status(400).json({ error: message });
    return;
  }
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Lỗi máy chủ.' });
});

// In production, serve the built client.
const clientDist = path.join(PROJECT_ROOT, 'client', 'dist');
if (existsSync(clientDist)) {
  app.use(express.static(clientDist));
  app.get('*', (_req, res) => {
    res.sendFile(path.join(clientDist, 'index.html'));
  });
}

app.listen(PORT, () => {
  console.log(`VBPL formatter server listening on http://localhost:${PORT}`);
  console.log(`Using Python: ${PYTHON_BIN}, script: ${SCRIPT_PATH}`);
});
