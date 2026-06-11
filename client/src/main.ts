const dropzone = document.getElementById('dropzone') as HTMLDivElement;
const fileInput = document.getElementById('file-input') as HTMLInputElement;
const fileInfo = document.getElementById('file-info') as HTMLDivElement;
const fileNameEl = document.getElementById('file-name') as HTMLSpanElement;
const clearBtn = document.getElementById('clear-btn') as HTMLButtonElement;
const formatBtn = document.getElementById('format-btn') as HTMLButtonElement;
const statusEl = document.getElementById('status') as HTMLDivElement;

let selectedFile: File | null = null;
let busy = false;

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

function selectFile(file: File | null) {
  if (file && !file.name.toLowerCase().endsWith('.docx')) {
    setStatus('error', 'Chỉ chấp nhận file Word (.docx).');
    return;
  }
  selectedFile = file;
  setStatus(null);
  if (file) {
    fileNameEl.textContent = `${file.name} (${formatSize(file.size)})`;
    fileInfo.hidden = false;
  } else {
    fileInfo.hidden = true;
    fileInput.value = '';
  }
  formatBtn.disabled = !file || busy;
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
  selectFile(fileInput.files?.[0] ?? null);
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
  const file = e.dataTransfer?.files?.[0] ?? null;
  if (file) selectFile(file);
});

clearBtn.addEventListener('click', () => selectFile(null));

// --- Submit ---

formatBtn.addEventListener('click', async () => {
  if (!selectedFile || busy) return;
  busy = true;
  formatBtn.disabled = true;
  formatBtn.textContent = 'Đang xử lý…';
  setStatus('info', 'Đang tải lên và định dạng văn bản…');

  try {
    const formData = new FormData();
    formData.append('file', selectedFile);

    const response = await fetch('/api/format', { method: 'POST', body: formData });

    if (!response.ok) {
      const message = await readErrorMessage(response);
      setStatus('error', message);
      return;
    }

    const blob = await response.blob();
    const downloadName =
      parseFilename(response.headers.get('Content-Disposition')) ??
      selectedFile.name.replace(/\.docx$/i, '') + '_formatted.docx';
    triggerDownload(blob, downloadName);
    setStatus('success', `Hoàn tất! Đã tải xuống "${downloadName}".`);
  } catch (err) {
    console.error(err);
    setStatus('error', 'Không kết nối được máy chủ. Vui lòng thử lại.');
  } finally {
    busy = false;
    formatBtn.textContent = 'Định dạng văn bản';
    formatBtn.disabled = !selectedFile;
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
