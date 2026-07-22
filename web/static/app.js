let sessionId = null;
let procesando = false;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('file-input');
const browseLink = document.getElementById('browse-link');
const fileInfo = document.getElementById('file-info');
const fileName = document.getElementById('file-name');
const changeFile = document.getElementById('change-file');
const logEl = document.getElementById('log');
const statusDot = document.getElementById('status-dot');
const statusText = document.getElementById('status-text');
const spinner = document.getElementById('spinner');
const downloadBtn = document.getElementById('btn-download');
const actions = document.querySelectorAll('.btn-action');
const btnTodo = document.getElementById('btn-todo');

const ACCIONES = ['numerar', 'insertar_seq', 'renumerar', 'verificar'];

browseLink.addEventListener('click', e => { e.preventDefault(); fileInput.click(); });
changeFile.addEventListener('click', () => {
  sessionId = null;
  fileInput.value = '';
  dropzone.hidden = false;
  fileInfo.hidden = true;
  setAccionesDisabled(true);
  downloadBtn.hidden = true;
  resetLog();
});

fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) subirArchivo(fileInput.files[0]);
});

dropzone.addEventListener('click', () => fileInput.click());

dropzone.addEventListener('dragover', e => {
  e.preventDefault();
  dropzone.classList.add('dragover');
});
dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('dragover');
});
dropzone.addEventListener('drop', e => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) subirArchivo(file);
});

ACCIONES.forEach(accion => {
  document.querySelector(`[data-action="${accion}"]`).addEventListener('click', () => {
    procesar(accion);
  });
});
btnTodo.addEventListener('click', () => procesar('todo'));

downloadBtn.addEventListener('click', () => {
  if (sessionId) {
    window.location.href = `/api/download/${sessionId}`;
  }
});

async function subirArchivo(file) {
  if (!file.name.endsWith('.docx')) {
    log('ERROR: Solo se permiten archivos .docx', 'error');
    return;
  }

  const form = new FormData();
  form.append('file', file);

  setStatus('busy', 'Subiendo...');
  log(`Subiendo archivo: ${file.name}...`);

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Error al subir');
    sessionId = data.session_id;

    dropzone.hidden = true;
    fileInfo.hidden = false;
    fileName.textContent = data.filename;
    setAccionesDisabled(false);
    setStatus('ready', 'Listo');
    log(`Archivo listo: ${data.filename}`, 'info');
  } catch (e) {
    setStatus('error', 'Error');
    log(`ERROR: ${e.message}`, 'error');
  }
}

async function procesar(accion) {
  if (procesando || !sessionId) return;
  procesando = true;

  const allBtns = [...document.querySelectorAll('.btn-action, #btn-todo, .btn-text')];
  allBtns.forEach(b => b.disabled = true);
  downloadBtn.hidden = true;
  setStatus('busy', 'Procesando...');
  spinner.hidden = false;

  const form = new FormData();
  form.append('session_id', sessionId);
  form.append('accion', accion);

  log(`\n${'='.repeat(48)}`, 'sep');
  log(`Iniciando: ${accion}...`);

  try {
    const res = await fetch('/api/procesar', {
      method: 'POST',
      body: form,
    });
    const data = await res.json();

    if (data.logs) {
      data.logs.forEach(msg => {
        if (msg.startsWith('ERROR')) log(msg, 'error');
        else if (msg.startsWith('  [ERROR]')) log(msg, 'error');
        else if (msg.startsWith('  [ADVERTENCIA]')) log(msg, 'warning');
        else log(msg);
      });
    }

    if (data.success) {
      setStatus('ready', 'Completado');
      log('Proceso completado exitosamente.', 'info');
      if (accion !== 'verificar') {
        downloadBtn.hidden = false;
      }
    } else {
      setStatus('error', 'Error');
      log(`\nERROR: ${data.error || data.resultado?.error || 'Error desconocido'}`, 'error');
    }
  } catch (e) {
    setStatus('error', 'Error');
    log(`\nERROR de conexión: ${e.message}`, 'error');
  } finally {
    procesando = false;
    spinner.hidden = true;
    if (sessionId) {
      setAccionesDisabled(false);
      allBtns.forEach(b => b.disabled = false);
      if (sessionId) document.getElementById('change-file').disabled = false;
    }
  }
}

function log(msg, type) {
  const line = document.createElement('div');
  line.className = 'log-line' + (type ? ' ' + type : '');
  line.textContent = msg;

  const empty = logEl.querySelector('.log-empty');
  if (empty) empty.remove();

  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function resetLog() {
  logEl.innerHTML = '<div class="log-empty">Sube un archivo .docx para comenzar.</div>';
}

function setAccionesDisabled(val) {
  ACCIONES.forEach(a => {
    document.querySelector(`[data-action="${a}"]`).disabled = val;
  });
  btnTodo.disabled = val;
}

function setStatus(state, text) {
  statusDot.className = 'status-dot' + (state === 'busy' ? ' busy' : state === 'error' ? ' error' : '');
  statusText.textContent = text;
}
