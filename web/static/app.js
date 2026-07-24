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
const progressTrack = document.getElementById('progress-track');

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
  setProgreso('reset');
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

function accionAFormData(accion) {
  const form = new FormData();
  form.append('session_id', sessionId);
  form.append('accion', accion);
  return form;
}

function mostrarLogs(data) {
  if (data.logs) {
    data.logs.forEach(msg => {
      if (msg.startsWith('ERROR') || msg.startsWith('  [ERROR]')) log(msg, 'error');
      else if (msg.startsWith('  [ADVERTENCIA]')) log(msg, 'warning');
      else log(msg);
    });
  }
}

async function llamarApi(accion) {
  const res = await fetch('/api/procesar', {
    method: 'POST',
    body: accionAFormData(accion),
  });
  const data = await res.json();
  mostrarLogs(data);
  if (!data.success) throw new Error(data.error || data.resultado?.error || 'Error desconocido');
  return data;
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

  log(`${'='.repeat(48)}`, 'sep');
  log(`Iniciando: ${NOMBRES[accion]}`);

  try {
    if (accion === 'todo') {
      setProgreso(1, 'active');
      const r1 = await llamarApi('numerar');
      setProgreso(1, 'completed');
      setProgreso(2, 'active');
      const r2 = await llamarApi('insertar_seq');
      setProgreso(2, 'completed');
      setProgreso(3, 'completed');
      setStatus('ready', 'Completado');
      log('Proceso completo finalizado exitosamente.', 'info');
      downloadBtn.hidden = false;
    } else if (accion === 'verificar') {
      setProgreso(3, 'active');
      const data = await llamarApi(accion);
      setProgreso(3, 'completed');
      setStatus('ready', 'Completado');
      log('Verificación completada.', 'info');
    } else {
      const pasos = { numerar: 1, insertar_seq: 2, renumerar: 1 };
      const p = pasos[accion] || 1;
      setProgreso(p, 'active');
      const data = await llamarApi(accion);
      setProgreso(p, 'completed');
      setProgreso(3, 'completed');
      setStatus('ready', 'Completado');
      log('Proceso completado exitosamente.', 'info');
      downloadBtn.hidden = false;
    }
  } catch (e) {
    setStatus('error', 'Error');
    log(`ERROR: ${e.message}`, 'error');
    setProgreso('reset');
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

function setProgreso(step, state) {
  for (let i = 1; i <= 3; i++) {
    const node = document.querySelector(`.step-node[data-step="${i}"]`);
    if (!node) continue;
    const circle = node.querySelector('.step-circle');
    const line = document.querySelector(`.step-line[data-from="${i}"]`);

    circle.innerHTML = '<span>' + i + '</span>';
    circle.className = 'step-circle';
    circle.style.background = '';
    circle.style.borderColor = '';
    circle.style.color = '';
    circle.style.boxShadow = '';
    if (line) {
      line.classList.remove('completed');
      line.style.background = '';
    }

    if (state === 'active' && i === step) {
      circle.style.borderColor = '#4ade80';
      circle.style.background = 'rgba(74,222,128,0.12)';
      circle.style.color = '#4ade80';
      circle.style.boxShadow = '0 0 0 4px rgba(74,222,128,0.12), 0 0 24px rgba(74,222,128,0.2)';
      circle.classList.add('active');
    } else if (i < step && state === 'active') {
      circle.innerHTML = '<span>✓</span>';
      circle.style.background = '#4ade80';
      circle.style.borderColor = '#4ade80';
      circle.style.color = '#0a0a14';
      circle.style.boxShadow = '0 0 0 4px rgba(74,222,128,0.15), 0 0 24px rgba(74,222,128,0.25)';
      circle.classList.add('completed');
      if (line) {
        line.style.background = '#4ade80';
        line.classList.add('completed');
      }
    } else if (i <= step && state === 'completed') {
      circle.innerHTML = '<span>✓</span>';
      circle.style.background = '#4ade80';
      circle.style.borderColor = '#4ade80';
      circle.style.color = '#0a0a14';
      circle.style.boxShadow = '0 0 0 4px rgba(74,222,128,0.15), 0 0 24px rgba(74,222,128,0.25)';
      circle.classList.add('completed');
      if (line) {
        line.style.background = '#4ade80';
        line.classList.add('completed');
      }
    }
  }
}

function setStatus(state, text) {
  statusDot.className = 'status-dot' + (state === 'busy' ? ' busy' : state === 'error' ? ' error' : '');
  statusText.textContent = text;
}
