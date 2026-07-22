let sessionId = null;
let procesando = false;
let actionActiva = null;

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
const stepIndicator = document.getElementById('step-indicator');
const downloadBtn = document.getElementById('btn-download');
const clearBtn = document.getElementById('btn-clear');
const actions = document.querySelectorAll('.btn-action');
const btnTodo = document.getElementById('btn-todo');

const ACCIONES = ['numerar', 'insertar_seq', 'renumerar', 'verificar'];
const NOMBRES = { numerar: 'Numerar', insertar_seq: 'Insertar SEQ', renumerar: 'Renumerar', verificar: 'Verificar', todo: 'Procesar todo' };

function ts() {
  return new Date().toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
}

browseLink.addEventListener('click', e => { e.preventDefault(); fileInput.click(); });
clearBtn.addEventListener('click', resetLog);
changeFile.addEventListener('click', () => {
  sessionId = null;
  fileInput.value = '';
  dropzone.hidden = false;
  fileInfo.hidden = true;
  downloadBtn.hidden = true;
  clearBtn.hidden = true;
  setAccionesDisabled(true);
  limpiarActiva();
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

function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1048576).toFixed(1) + ' MB';
}

async function subirArchivo(file) {
  if (!file.name.endsWith('.docx')) {
    log('ERROR: Solo se permiten archivos .docx', 'error');
    return;
  }

  const form = new FormData();
  form.append('file', file);

  setStatus('busy', 'Subiendo...');
  log(`Subiendo: ${file.name} (${formatSize(file.size)})...`);

  try {
    const res = await fetch('/api/upload', { method: 'POST', body: form });
    const data = await res.json();
    if (!res.ok) throw new Error(data.detail || 'Error al subir');
    sessionId = data.session_id;

    dropzone.hidden = true;
    fileInfo.hidden = false;
    fileName.textContent = `${data.filename} — ${formatSize(file.size)}`;
    clearBtn.hidden = false;
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
  actionActiva = accion;
  limpiarActiva();
  marcarActiva(accion);

  const allBtns = [...document.querySelectorAll('.btn-action, #btn-todo, .btn-text')];
  allBtns.forEach(b => b.disabled = true);
  downloadBtn.hidden = true;
  setStatus('busy', 'Procesando...');
  spinner.hidden = false;
  stepIndicator.hidden = false;
  stepIndicator.textContent = NOMBRES[accion] + '...';

  const form = new FormData();
  form.append('session_id', sessionId);
  form.append('accion', accion);

  log(`${'='.repeat(48)}`, 'sep');
  log(`Iniciando: ${NOMBRES[accion]}`);

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
      log(`ERROR: ${data.error || data.resultado?.error || 'Error desconocido'}`, 'error');
    }
  } catch (e) {
    setStatus('error', 'Error');
    log(`ERROR de conexión: ${e.message}`, 'error');
  } finally {
    procesando = false;
    spinner.hidden = true;
    stepIndicator.hidden = true;
    if (sessionId) {
      setAccionesDisabled(false);
      allBtns.forEach(b => b.disabled = false);
      document.getElementById('change-file').disabled = false;
    }
  }
}

function log(msg, type) {
  const line = document.createElement('div');
  line.className = 'log-line' + (type ? ' ' + type : '');

  const time = document.createElement('span');
  time.className = 'ts';
  time.textContent = ts();
  line.appendChild(time);

  const text = document.createElement('span');
  text.textContent = msg;
  line.appendChild(text);

  const empty = logEl.querySelector('.log-empty');
  if (empty) empty.remove();

  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}

function resetLog() {
  logEl.innerHTML = '<div class="log-empty">Sube un archivo .docx para comenzar.</div>';
  clearBtn.hidden = !sessionId;
}

function marcarActiva(accion) {
  const sel = accion === 'todo' ? '#btn-todo' : `[data-action="${accion}"]`;
  const el = document.querySelector(sel);
  if (el) el.classList.add('active');
}

function limpiarActiva() {
  document.querySelectorAll('.btn-action.active, #btn-todo.active').forEach(el => el.classList.remove('active'));
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
