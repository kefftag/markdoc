(() => {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('file-input');
  const fileInfo = document.getElementById('file-info');
  const fileName = document.getElementById('file-name');
  const convertBtn = document.getElementById('convert-btn');
  const status = document.getElementById('status');

  let selectedFile = null;

  // ── Drag and drop events ─────────────────────────────────────────────────
  ['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, e => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, e => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', e => {
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
  });

  // Keyboard accessibility for the drop zone
  dropZone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  // Click on drop zone triggers file browse (except when clicking the button directly)
  dropZone.addEventListener('click', e => {
    if (e.target !== fileInput && !e.target.closest('label')) {
      fileInput.click();
    }
  });

  // ── File input change ────────────────────────────────────────────────────
  fileInput.addEventListener('change', () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
    fileInput.value = ''; // reset so same file can be re-selected
  });

  // ── File handling ────────────────────────────────────────────────────────
  function handleFile(file) {
    const allowed = ['.md', '.markdown', '.txt'];
    const ext = file.name.slice(file.name.lastIndexOf('.')).toLowerCase();

    if (!allowed.includes(ext)) {
      setStatus('Only .md, .markdown, or .txt files are supported.', 'error');
      return;
    }

    selectedFile = file;
    fileName.textContent = file.name;
    fileInfo.hidden = false;
    setStatus('');
  }

  // ── Conversion ───────────────────────────────────────────────────────────
  convertBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    convertBtn.disabled = true;
    setStatus('Converting…', 'loading');

    try {
      const formData = new FormData();
      formData.append('file', selectedFile);

      const response = await fetch('/convert', { method: 'POST', body: formData });

      if (!response.ok) {
        const err = await response.json().catch(() => ({ error: 'Unknown error' }));
        throw new Error(err.error || `Server error ${response.status}`);
      }

      // Trigger download
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const baseName = selectedFile.name.slice(0, selectedFile.name.lastIndexOf('.'));
      a.href = url;
      a.download = `${baseName}.docx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus('Done! Your file is downloading.', 'success');
    } catch (err) {
      setStatus(`Error: ${err.message}`, 'error');
    } finally {
      convertBtn.disabled = false;
    }
  });

  // ── Helpers ──────────────────────────────────────────────────────────────
  function setStatus(message, type = '') {
    status.textContent = message;
    status.className = type;
  }
})();
