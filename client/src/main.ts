import { zip, type Zippable } from 'fflate';

const dropzone = document.getElementById('dropzone') as HTMLDivElement;
const fileInput = document.getElementById('file-input') as HTMLInputElement;
const fileList = document.getElementById('file-list') as HTMLUListElement;
const formatBtn = document.getElementById('format-btn') as HTMLButtonElement;
const statusEl = document.getElementById('status') as HTMLDivElement;
const downloadAllBtn = document.getElementById('download-all-btn') as HTMLButtonElement;
const resultList = document.getElementById('result-list') as HTMLUListElement;

// How many files to upload/process at once. Each goes to /api/format as its
// own request, so this just bounds concurrency against the serverless backend.
const MAX_CONCURRENCY = 4;

let selectedFiles: File[] = [];
let busy = false;
let zipping = false;
// Successful results of the last batch, with ZIP-unique entry names.
let zipEntries: { name: string; blob: Blob }[] = [];

function setStatus(kind: 'info' | 'success' | 'error' | null, message = '') {
  if (!kind) {
    statusEl.hidden = true;
    statusEl.textContent = '';
    statusEl.className = 'status';
    return;
  }
  statusEl.hidden = false;
  statusEl.textContent = message;
  statusEl.className = `status status--${kind}`;
}

function addFiles(files: File[]) {
  const docx = files.filter((f) => f.name.toLowerCase().endsWith('.docx'));
  const rejected = files.length - docx.length;

  // De-duplicate by name + size so re-selecting the same file is a no-op.
  for (const file of docx) {
    const dup = selectedFiles.some((f) => f.name === file.name && f.size === file.size);
    if (!dup) selectedFiles.push(file);
  }

  if (rejected > 0) {
    setStatus('error', `Đã bỏ qua ${rejected} file không phải định dạng .docx.`);
  } else {
    setStatus(null);
  }
  renderFileList();
}

function removeFile(index: number) {
  selectedFiles.splice(index, 1);
  renderFileList();
}

function renderFileList() {
  fileList.textContent = '';
  if (selectedFiles.length === 0) {
    fileList.hidden = true;
  } else {
    fileList.hidden = false;
    selectedFiles.forEach((file, index) => {
      const li = document.createElement('li');
      li.className = 'file-info';

      const name = document.createElement('span');
      name.className = 'file-name';
      name.textContent = `${file.name} (${formatSize(file.size)})`;

      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'clear-btn';
      remove.setAttribute('aria-label', `Bỏ chọn ${file.name}`);
      remove.textContent = '✕';
      remove.disabled = busy;
      remove.addEventListener('click', () => removeFile(index));

      li.append(name, remove);
      fileList.append(li);
    });
  }
  formatBtn.disabled = selectedFiles.length === 0 || busy;
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

// --- Dropzone interactions ---

dropzone.addEventListener('click', () => fileInput.click());
dropzone.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener('change', () => {
  addFiles(Array.from(fileInput.files ?? []));
});

for (const eventName of ['dragenter', 'dragover'] as const) {
  dropzone.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropzone.classList.add('dropzone--active');
  });
}
for (const eventName of ['dragleave', 'drop'] as const) {
  dropzone.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropzone.classList.remove('dropzone--active');
  });
}
dropzone.addEventListener('drop', (e) => {
  const files = Array.from(e.dataTransfer?.files ?? []);
  if (files.length) addFiles(files);
});

// --- Submit ---

interface FormatResult {
  file: File;
  ok: boolean;
  blob?: Blob;
  downloadName?: string;
  error?: string;
}

async function formatOne(file: File): Promise<FormatResult> {
  try {
    const formData = new FormData();
    formData.append('file', file);
    const response = await fetch('/api/format', { method: 'POST', body: formData });

    if (!response.ok) {
      return { file, ok: false, error: await readErrorMessage(response) };
    }

    const blob = await response.blob();
    const downloadName =
      parseFilename(response.headers.get('Content-Disposition')) ??
      file.name.replace(/\.docx$/i, '') + '_formatted.docx';
    return { file, ok: true, blob, downloadName };
  } catch {
    return { file, ok: false, error: 'Không kết nối được máy chủ.' };
  }
}

formatBtn.addEventListener('click', async () => {
  if (selectedFiles.length === 0 || busy) return;
  busy = true;
  formatBtn.disabled = true;
  formatBtn.textContent = 'Đang xử lý…';
  resultList.hidden = true;
  resultList.textContent = '';
  downloadAllBtn.hidden = true;
  zipEntries = [];
  renderFileList(); // disable the per-file remove buttons
  setStatus('info', `Đang tải lên và định dạng ${selectedFiles.length} văn bản…`);

  const queue = [...selectedFiles];
  const results: FormatResult[] = [];

  // Bounded-concurrency worker pool: each worker pulls from the shared queue.
  const worker = async () => {
    let file: File | undefined;
    while ((file = queue.shift())) {
      results.push(await formatOne(file));
      setStatus('info', `Đã xử lý ${results.length}/${selectedFiles.length} văn bản…`);
    }
  };
  await Promise.all(
    Array.from({ length: Math.min(MAX_CONCURRENCY, selectedFiles.length) }, worker),
  );

  renderResults(results);

  const succeeded = results.filter((r) => r.ok).length;
  const failed = results.length - succeeded;
  if (failed === 0) {
    setStatus('success', `Hoàn tất! Đã định dạng ${succeeded} văn bản. Bấm để tải xuống.`);
  } else if (succeeded === 0) {
    setStatus('error', `Không định dạng được văn bản nào (${failed} lỗi).`);
  } else {
    setStatus('info', `Hoàn tất ${succeeded} văn bản, ${failed} file lỗi. Xem chi tiết bên dưới.`);
  }

  busy = false;
  formatBtn.textContent = 'Xử lý văn bản';
  renderFileList();
});

function renderResults(results: FormatResult[]) {
  resultList.textContent = '';
  if (results.length === 0) {
    resultList.hidden = true;
    return;
  }
  resultList.hidden = false;

  // Collect successes for "Tải tất cả", de-duplicating ZIP entry names
  // (same-name uploads of different sizes produce the same downloadName).
  zipEntries = [];
  const usedNames = new Set<string>();
  for (const r of results) {
    if (!r.ok || !r.blob || !r.downloadName) continue;
    let name = r.downloadName;
    for (let n = 2; usedNames.has(name); n++) {
      name = r.downloadName.replace(/(\.docx)$/i, ` (${n})$1`);
    }
    usedNames.add(name);
    zipEntries.push({ name, blob: r.blob });
  }
  downloadAllBtn.hidden = zipEntries.length < 2;
  downloadAllBtn.textContent = `Tải tất cả (${zipEntries.length} file, .zip)`;

  for (const result of results) {
    const li = document.createElement('li');
    li.className = `result-item result-item--${result.ok ? 'ok' : 'error'}`;

    const name = document.createElement('span');
    name.className = 'result-name';
    name.textContent = result.file.name;

    li.append(name);

    if (result.ok && result.blob && result.downloadName) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'download-btn';
      button.textContent = 'Tải xuống';
      const blob = result.blob;
      const downloadName = result.downloadName;
      button.addEventListener('click', () => triggerDownload(blob, downloadName));
      li.append(button);
    } else {
      const error = document.createElement('span');
      error.className = 'result-error';
      error.textContent = result.error ?? 'Lỗi không xác định.';
      li.append(error);
    }

    resultList.append(li);
  }
}

// --- Download all as ZIP ---

function zipAsync(files: Zippable): Promise<Uint8Array<ArrayBuffer>> {
  return new Promise((resolve, reject) => {
    // level 0: .docx files are already deflate-compressed ZIP archives, so
    // re-compressing wastes CPU for ~0% size gain. Store them as-is.
    // The cast is sound: fflate allocates plain ArrayBuffers, but its types
    // predate TS 5.7's generic Uint8Array<ArrayBuffer> (needed for BlobPart).
    zip(files, { level: 0 }, (err, data) =>
      err ? reject(err) : resolve(data as Uint8Array<ArrayBuffer>),
    );
  });
}

downloadAllBtn.addEventListener('click', async () => {
  if (zipping || zipEntries.length === 0) return;
  zipping = true;
  downloadAllBtn.disabled = true;
  const label = downloadAllBtn.textContent;
  downloadAllBtn.textContent = 'Đang nén…';

  try {
    const files: Zippable = {};
    for (const entry of zipEntries) {
      files[entry.name] = new Uint8Array(await entry.blob.arrayBuffer());
    }
    const data = await zipAsync(files);
    triggerDownload(new Blob([data], { type: 'application/zip' }), 'van-ban-da-dinh-dang.zip');
    setStatus('success', `Đã tải xuống ${zipEntries.length} văn bản trong "van-ban-da-dinh-dang.zip".`);
  } catch (err) {
    console.error(err);
    setStatus('error', 'Không nén được file. Vui lòng tải từng file một.');
  } finally {
    zipping = false;
    downloadAllBtn.disabled = false;
    downloadAllBtn.textContent = label;
  }
});

async function readErrorMessage(response: Response): Promise<string> {
  // Vercel rejects bodies over 4.5 MB before they reach the API.
  if (response.status === 413) {
    return 'File quá lớn. Kích thước tối đa là 4,5 MB.';
  }
  try {
    const data = (await response.json()) as { error?: string };
    if (data.error) return data.error;
  } catch {
    // non-JSON error body
  }
  return `Xử lý thất bại (mã lỗi ${response.status}).`;
}

function parseFilename(contentDisposition: string | null): string | null {
  if (!contentDisposition) return null;
  // RFC 5987: filename*=UTF-8''<percent-encoded>
  const star = /filename\*=UTF-8''([^;]+)/i.exec(contentDisposition);
  if (star) {
    try {
      return decodeURIComponent(star[1]);
    } catch {
      // fall through to plain filename
    }
  }
  const plain = /filename="?([^";]+)"?/i.exec(contentDisposition);
  return plain ? plain[1] : null;
}

function triggerDownload(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}
