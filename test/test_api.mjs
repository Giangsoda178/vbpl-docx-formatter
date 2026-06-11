// End-to-end API tests using Node's built-in fetch/FormData.
import { readFile } from 'node:fs/promises';

const BASE = process.env.API_BASE ?? 'http://127.0.0.1:3001';
let failures = 0;

function check(label, cond, detail = '') {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${label}${detail ? ' — ' + detail : ''}`);
  if (!cond) failures++;
}

// 1. Valid upload with Vietnamese filename
{
  const buf = await readFile('test/sample_nghi_dinh.docx');
  const fd = new FormData();
  fd.append('file', new Blob([buf]), 'Nghị định 141.docx');
  const res = await fetch(`${BASE}/api/format`, { method: 'POST', body: fd });
  const cd = res.headers.get('content-disposition') ?? '';
  const body = await res.arrayBuffer();
  check('valid docx returns 200', res.status === 200, `status=${res.status}`);
  check('output is non-empty docx (PK zip magic)',
    body.byteLength > 1000 && new Uint8Array(body)[0] === 0x50 && new Uint8Array(body)[1] === 0x4b);
  check('Content-Disposition keeps Vietnamese name',
    decodeURIComponent(cd).includes('Nghị định 141_formatted.docx'), cd);
}

// 2. Structurally invalid docx (no header table) -> 422
{
  // A docx with no tables: reuse output of a plain document made on the fly is overkill;
  // upload a tiny zip that python-docx rejects -> script exits non-zero.
  const fd = new FormData();
  fd.append('file', new Blob([Buffer.from('not a real docx')]), 'fake.docx');
  const res = await fetch(`${BASE}/api/format`, { method: 'POST', body: fd });
  const data = await res.json();
  check('corrupt docx rejected with 422 and clean error',
    res.status === 422 && typeof data.error === 'string' &&
    data.error.length > 0 && !data.error.includes('Traceback'),
    `status=${res.status}, error=${data.error}`);
}

// 3. Wrong extension -> 400
{
  const fd = new FormData();
  fd.append('file', new Blob([Buffer.from('hello')]), 'notes.txt');
  const res = await fetch(`${BASE}/api/format`, { method: 'POST', body: fd });
  const data = await res.json();
  check('non-docx rejected with 400', res.status === 400, `status=${res.status}, error=${data.error}`);
}

// 4. No file -> 400
{
  const res = await fetch(`${BASE}/api/format`, { method: 'POST', body: new FormData() });
  check('missing file rejected with 400', res.status === 400, `status=${res.status}`);
}

process.exitCode = failures ? 1 : 0;
