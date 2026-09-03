// =========================================================================
// CONFIGURACIÓN DE LA FUENTE DE DATOS
// Reemplaza esta URL por el link "Raw" de tu archivo en GitHub, por ejemplo:
// https://raw.githubusercontent.com/tu-usuario/tu-repo/main/Historial_de_alertas.csv
// También funciona con un .xlsx si lo prefieres.
// =========================================================================
const DATA_SOURCE_URL = 'https://raw.githubusercontent.com/huamanijacob-cmyk/Historial-de-neveras/main/Historial_de_alertas.xlsx';

const fixMojibake = (s) => {
  if (typeof s !== 'string') return s;
  try {
    const fixed = decodeURIComponent(escape(s));
    // heuristic: only accept if it actually removed mojibake markers
    return /Ã|â€/.test(s) ? fixed : s;
  } catch(e){ return s; }
};

function parseFechaHora(input){
  if(input === null || input === undefined || input === '') return null;
  // Excel/SheetJS with cellDates:true can hand us a real Date object
  if(input instanceof Date){
    return isNaN(input.getTime()) ? null : input;
  }
  // Excel serial date number fallback (in case cellDates wasn't applied)
  if(typeof input === 'number'){
    const parsed = XLSX.SSF.parse_date_code(input);
    if(!parsed) return null;
    return new Date(parsed.y, parsed.m-1, parsed.d, parsed.H||0, parsed.M||0);
  }
  const s = String(input).trim();
  if(!s) return null;
  // dd/mm/yyyy hh:mm (formato peruano estándar)
  let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
  if(m){
    const [_, d, mo, y, h, mi] = m;
    return new Date(+y, +mo-1, +d, +h, +mi);
  }
  // yyyy-mm-dd hh:mm (formato ISO, típico de exports CSV)
  m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[T\s](\d{1,2}):(\d{2})/);
  if(m){
    const [_, y, mo, d, h, mi] = m;
    return new Date(+y, +mo-1, +d, +h, +mi);
  }
  return null;
}

function pad(n){ return n.toString().padStart(2,'0'); }
function fmtDate(d){ return d ? `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()}` : '–'; }
function fmtDateTime(d){ return d ? `${fmtDate(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}` : '–'; }
function toDateInputVal(d){ return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; }

let RAW = [];        // normalized rows
let EPISODES = [];   // paired desconexión->reconexión episodes
let loadedFileName = '';
let fechaArchivoDesc = null; // fecha real del archivo (Last-Modified de GitHub)

const ALERT_EVENTS = ['Desconexión por energía','Intermitencia','Fuera de zona'];

function normalizeRows(rawRecords){
  const rows = [];
  for(const r of rawRecords){
    const device = (r['Codigo Nevera']||'').toString().trim();
    if(!device) continue;
    const fecha = parseFechaHora(r['Fecha Hora']);
    if(!fecha) continue;
    let bateria = null;
    const bVal = r['Bateria'];
    if(bVal !== null && bVal !== undefined && bVal !== ''){
      if(typeof bVal === 'number'){
        // SheetJS puede devolver la fracción cruda (0.27) o ya el número de porcentaje (27)
        bateria = Math.round(bVal <= 1 ? bVal*100 : bVal);
      } else {
        const bm = bVal.toString().match(/(\d+(\.\d+)?)/);
        if(bm) bateria = Math.round(parseFloat(bm[1]));
      }
    }
    const energiaRaw = fixMojibake((r['Energia']||'').toString().trim());
    const energia = /desconect/i.test(energiaRaw) ? 'Desconectado' : (/conect/i.test(energiaRaw) ? 'Conectado' : null);
    rows.push({
      device,
      locacion: fixMojibake((r['Locacion']||'Sin distrito').toString().trim()) || 'Sin distrito',
      distribuidor: (r['Distribuidor']||'Sin distribuidor').toString().trim(),
      evento: fixMojibake((r['Evento']||'').toString().trim()),
      fecha,
      cliente: fixMojibake((r['Cliente']||'Sin cliente').toString().trim()),
      bateria,
      energia
    });
  }
  rows.sort((a,b)=>a.fecha-b.fecha);
  return rows;
}

// Reconstruye episodios de corte siguiendo la SECUENCIA ORDENADA DE EVENTOS por
// activo (misma lógica que al ordenar el Excel por placa y luego por fecha/hora:
// un evento ya representa el hecho consumado — cuando aparece "Fuera de zona" ya
// pasaron los metros/minutos del umbral, cuando aparece "Intermitencia" ya se
// detectó el patrón — así que no hay que corregir marcas de tiempo, solo encadenar
// la historia). Un incidente se ABRE con cualquiera de los 3 eventos "problema"
// (Fuera de zona, Intermitencia, Desconexión por energía) y se CIERRA con
// Reconexión/Conectado. Mientras sigan apareciendo eventos problema sin que
// llegue una reconexión, es el mismo incidente continuo.
const PROBLEM_EVENTS = ['Fuera de zona','Intermitencia','Desconexión por energía'];
const CLOSE_EVENTS = ['Reconexión','Conectado'];

function buildEpisodes(rows){
  const byDevice = {};
  for(const r of rows){
    (byDevice[r.device] = byDevice[r.device] || []).push(r);
  }
  const episodes = [];
  const maxDate = rows.length ? rows[rows.length-1].fecha : new Date();
  for(const device in byDevice){
    const evs = byDevice[device]; // ya viene ordenado por fecha
    let openStart = null;     // fila donde empezó el incidente actual
    let startUnknown = false; // el incidente ya venía abierto desde la primera fila que tenemos de este activo
    for(let i=0;i<evs.length;i++){
      const cur = evs[i];
      if(PROBLEM_EVENTS.includes(cur.evento) && openStart === null){
        openStart = cur;
        if(i===0) startUnknown = true;
      } else if(CLOSE_EVENTS.includes(cur.evento) && openStart !== null){
        const start = openStart.fecha;
        const end = cur.fecha;
        const hours = (end - start) / 3600000;
        episodes.push({
          device, cliente: openStart.cliente, locacion: openStart.locacion,
          distribuidor: openStart.distribuidor, start, end, hours,
          ongoing: false, startUnknown, startEvento: openStart.evento
        });
        openStart = null;
        startUnknown = false;
      }
    }
    if(openStart !== null){ // sigue con el incidente abierto al final de los datos disponibles
      const start = openStart.fecha;
      const hours = (maxDate - start) / 3600000;
      episodes.push({
        device, cliente: openStart.cliente, locacion: openStart.locacion,
        distribuidor: openStart.distribuidor, start, end: null, hours,
        ongoing: true, startUnknown, startEvento: openStart.evento
      });
    }
  }
  return episodes;
}

function durationBucket(h){
  if(h < 1) return '<1h';
  if(h < 3) return '1-3h';
  if(h < 6) return '3-6h';
  return '>6h';
}
const DUR_BUCKETS = ['<1h','1-3h','3-6h','>6h'];

function horarioBucket(d){
  const h = d.getHours();
  if(h < 6) return '12am-6am';
  if(h < 12) return '6am-12pm';
  if(h < 18) return '12pm-6pm';
  return '6pm-12am';
}
const HORARIO_BUCKETS = ['12am-6am','6am-12pm','12pm-6pm','6pm-12am'];

// ---------------- File loading ----------------
const fileInput = document.getElementById('fileInput');
const loadBtn = document.getElementById('loadBtn');
const emptyLoadBtn = document.getElementById('emptyLoadBtn');
const emptyState = document.getElementById('emptyState');
const dashboardEl = document.getElementById('desconexionesContent');
const loader = document.getElementById('loader');

// El botón principal y el de reintento ahora vuelven a traer el archivo del
// repositorio en lugar de abrir el explorador de archivos.
loadBtn.onclick = () => fetchAndLoadData(true);
emptyLoadBtn.onclick = () => fetchAndLoadData(true);
// Se deja la carga manual solo como respaldo, oculta detrás de un enlace pequeño.
document.getElementById('manualFallbackLink').onclick = () => fileInput.click();
fileInput.onchange = (e) => { if(e.target.files[0]) handleFile(e.target.files[0]); };

['dragover','dragleave','drop'].forEach(evt=>{
  emptyState.addEventListener(evt, (e)=>{ e.preventDefault();
    if(evt==='dragover') emptyState.classList.add('drag');
    if(evt==='dragleave') emptyState.classList.remove('drag');
    if(evt==='drop'){ emptyState.classList.remove('drag'); if(e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }
  });
});

// Trae el archivo directamente desde DATA_SOURCE_URL (por ejemplo, un raw.githubusercontent.com)
// y lo procesa igual que si el usuario lo hubiera subido a mano.
// ---------------- Barra de estado de fuentes de datos ----------------
function fmtDateCorta(d){
  if(!d || isNaN(d.getTime())) return '';
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear().toString().slice(-2)}`;
}

// GitHub expone la fecha real del archivo (cuándo se subió/actualizó) en la
// cabecera 'Last-Modified' de la respuesta — no hay que adivinarla del contenido.
function getFechaArchivo(res){
  const header = res.headers.get('Last-Modified');
  if(!header) return null;
  const d = new Date(header);
  return isNaN(d.getTime()) ? null : d;
}

function setSourceStatus(which, state, text){
  // which: 'Desc' | 'Censo' — state: 'loading' | 'ok' | 'error'
  const dot = document.getElementById('sourceDot' + which);
  const status = document.getElementById('sourceStatus' + which);
  if(!dot || !status) return;
  dot.className = 'source-dot ' + state;
  status.textContent = text;
}

async function fetchAndLoadData(isRetry){
  loader.style.display = 'flex';
  document.getElementById('loaderText').textContent = isRetry ? 'Actualizando datos…' : 'Cargando datos…';
  emptyLoadBtn.style.display = 'none';
  document.getElementById('emptyTitle').textContent = isRetry ? 'Actualizando datos…' : 'Cargando datos…';
  document.getElementById('emptyText').textContent = 'Obteniendo el archivo más reciente desde el repositorio.';
  setSourceStatus('Desc', 'loading', 'cargando…');

  try{
    // cachebust evita que el navegador muestre una copia vieja en caché
    const sep = DATA_SOURCE_URL.includes('?') ? '&' : '?';
    const res = await fetch(DATA_SOURCE_URL + sep + 'cachebust=' + Date.now());
    if(!res.ok) throw new Error('No se pudo descargar el archivo (HTTP ' + res.status + ')');
    fechaArchivoDesc = getFechaArchivo(res);

    const cleanUrl = DATA_SOURCE_URL.split('?')[0];
    const ext = cleanUrl.split('.').pop().toLowerCase();
    loadedFileName = cleanUrl.split('/').pop();

    if(ext === 'csv'){
      const text = await res.text();
      const parsed = Papa.parse(text, { header:true, skipEmptyLines:true });
      onDataParsed(parsed.data);
    } else if(ext === 'xlsx' || ext === 'xls'){
      const buf = await res.arrayBuffer();
      const wb = XLSX.read(buf, {type:'array', cellDates:true});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, {defval:'', cellDates:true});
      onDataParsed(data);
    } else {
      throw new Error('Extensión de archivo no reconocida (usa .csv o .xlsx)');
    }
  } catch(err){
    loader.style.display = 'none';
    document.getElementById('emptyTitle').textContent = 'No se pudo cargar el archivo automáticamente';
    document.getElementById('emptyText').textContent = 'Detalle: ' + err.message + '. Revisa que la URL del repositorio sea correcta, pública y que el enlace apunte al archivo "Raw".';
    emptyLoadBtn.textContent = '↻ Reintentar';
    emptyLoadBtn.style.display = 'inline-flex';
    dashboardEl.style.display = 'none';
    emptyState.style.display = 'flex';
    setSourceStatus('Desc', 'error', 'error al cargar');
    console.error('Error cargando Desconexiones y Alertas:', err);
  }
}

function handleFile(file){
  loader.style.display = 'flex';
  document.getElementById('loaderText').textContent = 'Procesando archivo…';
  loadedFileName = file.name;
  const ext = file.name.split('.').pop().toLowerCase();
  if(ext === 'csv'){
    Papa.parse(file, {
      header:true, skipEmptyLines:true, encoding:'UTF-8',
      complete: (res) => onDataParsed(res.data),
      error: (err) => { alert('No se pudo leer el CSV: '+err); loader.style.display='none'; }
    });
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, {type:'array', cellDates:true});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(sheet, {defval:'', cellDates:true});
      onDataParsed(data);
    };
    reader.readAsArrayBuffer(file);
  }
}

function onDataParsed(records){
  setTimeout(()=>{ // let loader paint
    RAW = normalizeRows(records);
    EPISODES = buildEpisodes(RAW);
    if(RAW.length === 0){
      document.getElementById('emptyTitle').textContent = 'El archivo no tiene filas válidas';
      document.getElementById('emptyText').textContent = 'Revisa el formato de columnas (se esperan encabezados como "Codigo Nevera", "Fecha Hora", "Evento"...).';
      emptyLoadBtn.textContent = '↻ Reintentar';
      emptyLoadBtn.style.display = 'inline-flex';
      dashboardEl.style.display = 'none';
      emptyState.style.display = 'flex';
      loader.style.display='none';
      return;
    }
    initFiltersUI();
    initPlacaUI();
    applyFiltersAndRender();
    emptyState.style.display = 'none';
    dashboardEl.style.display = 'block';
    document.getElementById('liveDot').style.display = 'inline-block';
    loader.style.display = 'none';
    setSourceStatus('Desc', 'ok', fechaArchivoDesc ? `archivo del ${fmtDateCorta(fechaArchivoDesc)}` : `cargado · ${RAW.length.toLocaleString('es-PE')} filas`);
  }, 30);
}

// ---------------- Componente: selector múltiple con búsqueda (estilo Excel) ----------------
// Comportamiento: "Todos" es el estado por defecto (checkbox marcado = sin filtro).
// Si desmarcas un elemento individual, "Todos" se desmarca automáticamente y queda
// aplicado el filtro explícito (todo menos lo que quitaste). Nunca se permite llegar
// a 0 elementos marcados: si desmarcas el último que quedaba, vuelve a "Todos".
function createMultiSelect(containerId, {placeholder='Todos', onChange=()=>{}} = {}){
  const container = document.getElementById(containerId);
  if(!container) return null;
  container.innerHTML = `
    <button type="button" class="msel-control">
      <span class="msel-summary">${placeholder}</span>
      <span class="msel-caret">▾</span>
    </button>
    <div class="msel-panel">
      <input type="text" class="msel-search" placeholder="Buscar...">
      <label class="msel-todos-row">
        <input type="checkbox" class="msel-todos-checkbox" checked>
        <span>Todos</span>
      </label>
      <div class="msel-divider"></div>
      <div class="msel-options"></div>
    </div>
  `;
  const control = container.querySelector('.msel-control');
  const summary = container.querySelector('.msel-summary');
  const search = container.querySelector('.msel-search');
  const todosCheckbox = container.querySelector('.msel-todos-checkbox');
  const optionsWrap = container.querySelector('.msel-options');

  let allOptions = [];
  let selected = new Set(); // solo tiene sentido cuando todosMode === false
  let todosMode = true;

  function updateSummary(){
    if(todosMode) summary.textContent = placeholder;
    else if(selected.size === 0) summary.textContent = 'Ninguno seleccionado';
    else if(selected.size === 1) summary.textContent = [...selected][0];
    else summary.textContent = `${selected.size} seleccionados`;
    control.classList.toggle('has-value', !todosMode);
    todosCheckbox.checked = todosMode;
  }

  function isChecked(opt){ return todosMode ? true : selected.has(opt); }

  function renderOptions(filterText){
    const ft = (filterText||'').trim().toLowerCase();
    const visible = ft ? allOptions.filter(o=>o.toLowerCase().includes(ft)) : allOptions;
    optionsWrap.innerHTML = visible.map(o=>`
      <label class="msel-option">
        <input type="checkbox" value="${o.replace(/"/g,'&quot;')}" ${isChecked(o) ? 'checked' : ''}>
        <span>${o}</span>
      </label>
    `).join('') || '<div class="msel-empty">Sin resultados</div>';
  }

  function handleIndividualToggle(v, checked){
    if(todosMode){
      // salir de "Todos": queda todo marcado excepto el que se acaba de desmarcar
      // (si en cambio lo que hizo fue marcar uno estando ya todo marcado, no hay nada que hacer)
      if(!checked){
        todosMode = false;
        selected = new Set(allOptions);
        selected.delete(v);
      }
    } else {
      if(checked) selected.add(v); else selected.delete(v);
      // Si terminan marcados todos de nuevo, se colapsa a "Todos" (más limpio).
      // Ya NO forzamos volver a "Todos" al llegar a 0 — estilo Excel, se permite
      // quedar en "ninguno seleccionado" a propósito (el filtro excluye todo el campo).
      if(selected.size === allOptions.length){
        todosMode = true;
        selected.clear();
      }
    }
  }

  optionsWrap.addEventListener('change', (e)=>{
    if(e.target.matches('input[type=checkbox]')){
      handleIndividualToggle(e.target.value, e.target.checked);
      renderOptions(search.value);
      updateSummary();
      onChange();
    }
  });

  todosCheckbox.addEventListener('change', ()=>{
    // Estilo Excel: el checkbox "Todos" es un interruptor real.
    // Marcarlo = seleccionar todo. Desmarcarlo = deseleccionar todo (el filtro
    // pasa a excluir el campo por completo, no a "mostrar todo").
    todosMode = todosCheckbox.checked;
    selected.clear();
    renderOptions(search.value);
    updateSummary();
    onChange();
  });

  search.addEventListener('input', ()=> renderOptions(search.value));

  control.addEventListener('click', (e)=>{
    e.stopPropagation();
    const isOpen = container.classList.contains('open');
    document.querySelectorAll('.msel.open, .drange.open').forEach(m=>m.classList.remove('open'));
    if(!isOpen){
      container.classList.add('open');
      search.value = '';
      renderOptions();
      search.focus();
    }
  });
  document.addEventListener('click', (e)=>{
    if(!container.contains(e.target)) container.classList.remove('open');
  });

  return {
    setOptions(newOptions){
      allOptions = newOptions;
      if(!todosMode){
        selected = new Set([...selected].filter(v=>allOptions.includes(v))); // descarta selección que ya no existe
      }
      renderOptions(search.value);
      updateSummary();
    },
    // null = sin filtro (equivale a "Todos"). Array (incluso vacío) = filtro
    // explícito; un array vacío significa "no debe coincidir nada" (estilo Excel).
    getValues(){ return todosMode ? null : [...selected]; },
    // Selecciona un único valor (usado al hacer clic en una barra de "Avance por...")
    selectOnly(value){
      todosMode = false;
      selected = new Set([value]);
      renderOptions(search.value);
      updateSummary();
    },
    clear(){ todosMode = true; selected.clear(); renderOptions(search.value); updateSummary(); }
  };
}

// ---------------- Componente: selector de rango de fechas (calendario propio) ----------------
function createDateRangePicker(containerId, {minDate, maxDate, onChange}){
  const container = document.getElementById(containerId);
  if(!container) return null;
  const norm = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  minDate = norm(minDate); maxDate = norm(maxDate);

  container.innerHTML = `
    <button type="button" class="drange-control">
      <span class="drange-summary"></span>
      <span class="drange-icon">▾</span>
    </button>
    <div class="drange-panel">
      <div class="drange-presets">
        <span class="drange-preset" data-preset="all">Todo el rango</span>
        <span class="drange-preset" data-preset="7">Últimos 7 días</span>
        <span class="drange-preset" data-preset="30">Últimos 30 días</span>
      </div>
      <div class="drange-month-nav">
        <span class="drange-nav-btn" data-nav="-1">‹</span>
        <span class="drange-month-label"></span>
        <span class="drange-nav-btn" data-nav="1">›</span>
      </div>
      <div class="drange-weekdays"><span>D</span><span>L</span><span>M</span><span>M</span><span>J</span><span>V</span><span>S</span></div>
      <div class="drange-days"></div>
      <div class="drange-footer">
        <span class="drange-hint">Elige la fecha de inicio</span>
        <button type="button" class="drange-apply">Listo</button>
      </div>
    </div>
  `;
  const control = container.querySelector('.drange-control');
  const summary = container.querySelector('.drange-summary');
  const monthLabel = container.querySelector('.drange-month-label');
  const daysWrap = container.querySelector('.drange-days');
  const hint = container.querySelector('.drange-hint');
  const prevBtn = container.querySelector('[data-nav="-1"]');
  const nextBtn = container.querySelector('[data-nav="1"]');
  const MONTHS = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  let rangeStart = minDate, rangeEnd = maxDate;
  let pickStage = 'start'; // 'start' | 'end' — en qué punto de la selección de 2 clics estamos
  let tempStart = rangeStart, tempEnd = rangeEnd;
  let viewMonth = new Date(maxDate.getFullYear(), maxDate.getMonth(), 1);

  function updateSummary(){
    const full = rangeStart.getTime()===minDate.getTime() && rangeEnd.getTime()===maxDate.getTime();
    summary.textContent = full ? 'Todo el rango' : `${fmtDate(rangeStart)} → ${fmtDate(rangeEnd)}`;
    control.classList.toggle('has-value', !full);
  }

  function renderCalendar(){
    monthLabel.textContent = `${MONTHS[viewMonth.getMonth()]} ${viewMonth.getFullYear()}`;
    const firstDow = new Date(viewMonth.getFullYear(), viewMonth.getMonth(), 1).getDay();
    const daysInMonth = new Date(viewMonth.getFullYear(), viewMonth.getMonth()+1, 0).getDate();
    let html = '';
    for(let i=0;i<firstDow;i++) html += `<span class="drange-day empty"></span>`;
    for(let day=1; day<=daysInMonth; day++){
      const d = new Date(viewMonth.getFullYear(), viewMonth.getMonth(), day);
      const disabled = d < minDate || d > maxDate;
      let cls = 'drange-day';
      if(disabled){
        cls += ' disabled';
      } else {
        if(d >= tempStart && d <= tempEnd) cls += ' in-range';
        if(d.getTime()===tempStart.getTime()) cls += ' range-start';
        if(d.getTime()===tempEnd.getTime()) cls += ' range-end';
      }
      html += `<span class="${cls}" data-day="${disabled?'':toDateInputVal(d)}">${day}</span>`;
    }
    daysWrap.innerHTML = html;
    prevBtn.classList.toggle('disabled', viewMonth.getFullYear()===minDate.getFullYear() && viewMonth.getMonth()===minDate.getMonth());
    nextBtn.classList.toggle('disabled', viewMonth.getFullYear()===maxDate.getFullYear() && viewMonth.getMonth()===maxDate.getMonth());
  }

  daysWrap.addEventListener('click', (e)=>{
    const el = e.target.closest('.drange-day');
    if(!el || el.classList.contains('disabled') || el.classList.contains('empty')) return;
    const d = new Date(el.dataset.day + 'T00:00:00');
    if(pickStage === 'start'){
      tempStart = d; tempEnd = d;
      pickStage = 'end';
      hint.textContent = 'Elige la fecha de fin';
    } else {
      if(d < tempStart){ tempEnd = tempStart; tempStart = d; } else { tempEnd = d; }
      pickStage = 'start';
      hint.textContent = 'Elige la fecha de inicio';
      rangeStart = tempStart; rangeEnd = tempEnd;
      updateSummary();
      onChange();
    }
    renderCalendar();
  });

  prevBtn.addEventListener('click', ()=>{
    if(prevBtn.classList.contains('disabled')) return;
    viewMonth = new Date(viewMonth.getFullYear(), viewMonth.getMonth()-1, 1);
    renderCalendar();
  });
  nextBtn.addEventListener('click', ()=>{
    if(nextBtn.classList.contains('disabled')) return;
    viewMonth = new Date(viewMonth.getFullYear(), viewMonth.getMonth()+1, 1);
    renderCalendar();
  });

  container.querySelectorAll('.drange-preset').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const p = btn.dataset.preset;
      if(p === 'all'){
        tempStart = minDate; tempEnd = maxDate;
      } else {
        const days = parseInt(p, 10);
        tempEnd = maxDate;
        tempStart = new Date(maxDate);
        tempStart.setDate(tempStart.getDate() - (days - 1));
        if(tempStart < minDate) tempStart = minDate;
      }
      rangeStart = tempStart; rangeEnd = tempEnd;
      pickStage = 'start';
      hint.textContent = 'Elige la fecha de inicio';
      viewMonth = new Date(tempEnd.getFullYear(), tempEnd.getMonth(), 1);
      renderCalendar();
      updateSummary();
      onChange();
    });
  });

  container.querySelector('.drange-apply').addEventListener('click', ()=> container.classList.remove('open'));

  control.addEventListener('click', (e)=>{
    e.stopPropagation();
    const isOpen = container.classList.contains('open');
    document.querySelectorAll('.msel.open, .drange.open').forEach(m=>m.classList.remove('open'));
    if(!isOpen){
      container.classList.add('open');
      tempStart = rangeStart; tempEnd = rangeEnd; pickStage = 'start';
      hint.textContent = 'Elige la fecha de inicio';
      viewMonth = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), 1);
      renderCalendar();
    }
  });
  document.addEventListener('click', (e)=>{
    if(!container.contains(e.target)) container.classList.remove('open');
  });

  updateSummary();

  return {
    getRange(){
      return {
        desde: rangeStart,
        hasta: new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), rangeEnd.getDate(), 23, 59, 59)
      };
    },
    reset(){
      rangeStart = minDate; rangeEnd = maxDate; tempStart = minDate; tempEnd = maxDate;
      updateSummary();
    }
  };
}

// ---------------- Filters ----------------
function uniqueSorted(arr){ return [...new Set(arr)].sort((a,b)=>a.localeCompare(b,'es')); }

let msDistribuidor, msDistrito, msEvento, drFecha;

function initFiltersUI(){
  const distribuidores = uniqueSorted(RAW.map(r=>r.distribuidor));
  const distritos = uniqueSorted(RAW.map(r=>r.locacion));
  const eventos = uniqueSorted(RAW.map(r=>r.evento));
  const maxDate = RAW[RAW.length-1].fecha;
  const minDate = RAW[0].fecha;

  if(!msDistribuidor) msDistribuidor = createMultiSelect('fDistribuidor', {placeholder:'Todos', onChange: applyFiltersAndRender});
  if(!msDistrito) msDistrito = createMultiSelect('fDistrito', {placeholder:'Todos', onChange: applyFiltersAndRender});
  if(!msEvento) msEvento = createMultiSelect('fEvento', {placeholder:'Todos', onChange: applyFiltersAndRender});
  msDistribuidor.setOptions(distribuidores);
  msDistrito.setOptions(distritos);
  msEvento.setOptions(eventos);

  // El selector de fechas se reconstruye cada vez que cargan datos nuevos, porque
  // su rango mínimo/máximo depende del archivo cargado.
  drFecha = createDateRangePicker('fFechaRango', {minDate, maxDate, onChange: applyFiltersAndRender});

  ['fSearch','fOnlyOfficial'].forEach(id=>{
    document.getElementById(id).addEventListener('input', applyFiltersAndRender);
  });
  document.getElementById('clearFilters').onclick = () => {
    msDistribuidor.clear();
    msDistrito.clear();
    msEvento.clear();
    drFecha.reset();
    document.getElementById('fSearch').value = '';
    document.getElementById('fOnlyOfficial').checked = false;
    applyFiltersAndRender();
  };
}

function getFilters(){
  return {
    distribuidor: msDistribuidor ? msDistribuidor.getValues() : null,
    distrito: msDistrito ? msDistrito.getValues() : null,
    evento: msEvento ? msEvento.getValues() : null,
    desde: drFecha ? drFecha.getRange().desde : null,
    hasta: drFecha ? drFecha.getRange().hasta : null,
    search: document.getElementById('fSearch').value.trim().toLowerCase(),
    onlyOfficial: document.getElementById('fOnlyOfficial').checked
  };
}

function matchesRow(r, f){
  if(f.distribuidor !== null && !f.distribuidor.includes(r.distribuidor)) return false;
  if(f.distrito !== null && !f.distrito.includes(r.locacion)) return false;
  if(f.evento !== null && !f.evento.includes(r.evento)) return false;
  if(f.desde && r.fecha < f.desde) return false;
  if(f.hasta && r.fecha > f.hasta) return false;
  if(f.search){
    const s = f.search;
    if(!r.cliente.toLowerCase().includes(s) && !r.device.toLowerCase().includes(s)) return false;
  }
  return true;
}
function matchesEpisode(ep, f){
  if(f.distribuidor !== null && !f.distribuidor.includes(ep.distribuidor)) return false;
  if(f.distrito !== null && !f.distrito.includes(ep.locacion)) return false;
  if(f.desde && ep.start < f.desde) return false;
  if(f.hasta && ep.start > f.hasta) return false;
  if(f.onlyOfficial && ep.hours < 1) return false;
  if(f.search){
    const s = f.search;
    if(!ep.cliente.toLowerCase().includes(s) && !ep.device.toLowerCase().includes(s)) return false;
  }
  return true;
}

function isNocturnal(ep){
  const h = ep.start.getHours();
  return (h >= 21 || h < 6) && ep.hours >= 2;
}

// ---------------- Render ----------------
function applyFiltersAndRender(){
  const f = getFilters();
  const rows = RAW.filter(r=>matchesRow(r,f));
  const episodes = EPISODES.filter(e=>matchesEpisode(e,f));
  const maxDateAll = RAW[RAW.length-1].fecha;

  document.getElementById('lastUpdateText').textContent = `Última fecha de actualización: ${fmtDateTime(maxDateAll)}`;

  renderKPIs(rows, episodes);
  renderTop10Activos24h(rows, maxDateAll);
  renderAlertMatrix(rows);
  renderHorarioMatrix(episodes);
  renderDistritoMatrix(episodes);
  renderTop10Clientes6h(episodes);
}

function renderKPIs(rows, episodes){
  const alertRows = rows.filter(r=>ALERT_EVENTS.includes(r.evento));
  document.getElementById('kpiTotal').textContent = rows.length.toLocaleString('es-PE');
  document.getElementById('kpiTotalSub').textContent = `de ${RAW.length.toLocaleString('es-PE')} totales`;
  document.getElementById('kpiActivos').textContent = new Set(alertRows.map(r=>r.device)).size.toLocaleString('es-PE');
  document.getElementById('kpiClientes').textContent = new Set(alertRows.map(r=>r.cliente)).size.toLocaleString('es-PE');
  document.getElementById('kpiLargas').textContent = episodes.filter(e=>e.hours>=6).length.toLocaleString('es-PE');
  document.getElementById('kpiNocturnas').textContent = episodes.filter(isNocturnal).length.toLocaleString('es-PE');
  const lastBatteryByDevice = {};
  rows.forEach(r=>{ if(r.bateria !== null) lastBatteryByDevice[r.device] = r.bateria; });
  const critBat = Object.values(lastBatteryByDevice).filter(b=>b<20).length;
  document.getElementById('kpiBateria').textContent = critBat.toLocaleString('es-PE');
}

function heatColor(value, max){
  if(max<=0 || value<=0) return {bg:'transparent', fg:'inherit'};
  const t = Math.min(value/max, 1);
  // verde (bajo) -> dorado (medio) -> rojo (alto) — mismo tono "semáforo" que el
  // módulo de Censo, para que ambos se sientan parte de un mismo producto.
  let bg;
  if(t < 0.5){
    const p = t/0.5;
    bg = mix('#d9f0df','#f2c94c',p);
  } else {
    const p = (t-0.5)/0.5;
    bg = mix('#f2c94c','#c0392b',p);
  }
  const fg = t > 0.62 ? '#ffffff' : '#1b2436';
  return {bg, fg};
}

// Celda de las tablas de calor como una "burbuja" circular de tamaño proporcional
// al valor, en vez de una celda rectangular sólida — se ve mucho menos a hoja de
// cálculo y más a un gráfico de datos moderno, manteniendo el mismo color semáforo.
function heatDot(value, max){
  if(max<=0 || value<=0) return `<td><span class="heat-dot-empty">–</span></td>`;
  const t = Math.min(value/max, 1);
  const size = Math.round(17 + t*23); // 17px .. 40px
  const hc = heatColor(value, max);
  const fontSize = size >= 27 ? 12 : 10.5;
  return `<td><span class="heat-dot" style="width:${size}px; height:${size}px; background:${hc.bg}; color:${hc.fg}; font-size:${fontSize}px;">${value}</span></td>`;
}

function mix(c1,c2,p){
  const a = hexToRgb(c1), b = hexToRgb(c2);
  const r = Math.round(a[0]+(b[0]-a[0])*p);
  const g = Math.round(a[1]+(b[1]-a[1])*p);
  const bl = Math.round(a[2]+(b[2]-a[2])*p);
  return `rgb(${r},${g},${bl})`;
}
function hexToRgb(h){ const n = parseInt(h.slice(1),16); return [(n>>16)&255,(n>>8)&255,n&255]; }

// ---------------- Gráficos de barra (Chart.js) para los Top 10 ----------------
const chartInstances = {};
function renderBarChart(containerId, items, maxVal){
  const container = document.getElementById(containerId);
  if(items.length === 0){
    if(chartInstances[containerId]){ chartInstances[containerId].destroy(); delete chartInstances[containerId]; }
    container.innerHTML = '<div style="color:var(--muted2); font-size:12.5px; padding:10px 0;">Sin datos en el rango filtrado.</div>';
    return;
  }
  if(!container.querySelector('canvas')){
    container.innerHTML = '<div style="position:relative;"><canvas></canvas></div>';
  }
  container.style.height = Math.max(180, items.length*32 + 26) + 'px';
  container.firstElementChild.style.height = container.style.height;
  const canvas = container.querySelector('canvas');
  const labels = items.map(i=>i.label);
  const data = items.map(i=>i.value);
  const colors = items.map(i=> maxVal ? mix('#2e9e4f','#c0392b', Math.min(i.value/maxVal,1)) : '#0c2461');

  if(chartInstances[containerId]){
    const ch = chartInstances[containerId];
    ch.data.labels = labels;
    ch.data.datasets[0].data = data;
    ch.data.datasets[0].backgroundColor = colors;
    ch.update();
  } else {
    chartInstances[containerId] = new Chart(canvas, {
      type: 'bar',
      data: { labels, datasets: [{ data, backgroundColor: colors, borderRadius: 6, barPercentage: 0.68, categoryPercentage: 0.85 }] },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 280 },
        plugins: {
          legend: { display: false },
          tooltip: {
            backgroundColor: '#0c2461', titleFont: { family: 'Inter', weight: '600' },
            bodyFont: { family: 'IBM Plex Mono' }, padding: 10, cornerRadius: 8, displayColors: false
          }
        },
        scales: {
          x: { beginAtZero: true, ticks: { precision: 0, font: { family: 'IBM Plex Mono', size: 11 }, color: '#6b7688' }, grid: { color: '#eef1f6' }, border: { display: false } },
          y: { ticks: { font: { family: 'Inter', size: 11.5 }, color: '#1b2436' }, grid: { display: false }, border: { display: false } }
        }
      }
    });
  }
}

function renderTop10Activos24h(rows, maxDateAll){
  const from = new Date(maxDateAll.getTime() - 24*3600000);
  const alertRows = rows.filter(r=>ALERT_EVENTS.includes(r.evento) && r.fecha >= from && r.fecha <= maxDateAll);
  const counts = {};
  alertRows.forEach(r=> counts[r.device] = (counts[r.device]||0)+1);
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([label,value])=>({label,value}));
  document.getElementById('badge24h').textContent = `${fmtDateTime(from)} → ${fmtDateTime(maxDateAll)}`;
  renderBarChart('top10Activos', top, top[0]?.value||0);
}

function extractClientCode(cliente){
  if(!cliente) return 'Sin código';
  const m = cliente.trim().match(/^(\S+)/);
  return m ? m[1] : cliente;
}

// Excel a veces convierte códigos largos a notación científica (ej. "2.203E+13"),
// lo que pierde dígitos y hace imposible reconstruir el código real.
const SCI_CODE_RE = /^\d(\.\d+)?E\+\d+$/i;
function isSciCode(code){ return SCI_CODE_RE.test(code); }

// Un código en notación científica ("2.9001E+13") conserva solo unos pocos dígitos
// significativos del código real. Si, redondeado a esa misma cantidad de dígitos,
// coincide con un código completo real, es matemáticamente el mismo cliente.
function sciMatchesCode(sciStr, fullCodeStr){
  const sciNum = parseFloat(sciStr);
  const fullNum = parseFloat(fullCodeStr);
  if(!isFinite(sciNum) || !isFinite(fullNum) || fullNum===0) return false;
  const m = sciStr.match(/^(\d)(\.(\d+))?E\+(\d+)/i);
  if(!m) return false;
  const sigDigits = 1 + (m[3] ? m[3].length : 0);
  const rounded = Number(fullNum.toPrecision(sigDigits));
  return Math.abs(rounded - sciNum) < 1;
}

// Intenta resolver los códigos en notación científica de UNA placa contra los
// códigos completos reales que esa misma placa mostró en otras filas. Solo se
// resuelve si hay una única coincidencia posible (evita falsos traspasos).
function buildCodeResolver(deviceRows){
  const knownCodes = new Set();
  deviceRows.forEach(r=>{
    const raw = extractClientCode(r.cliente);
    if(!isSciCode(raw)) knownCodes.add(raw);
  });
  const knownArr = [...knownCodes];
  const cache = {};
  return function resolve(raw){
    if(!isSciCode(raw)) return raw;
    if(cache[raw] !== undefined) return cache[raw];
    const matches = knownArr.filter(code=>sciMatchesCode(raw, code));
    const result = matches.length === 1 ? matches[0] : raw;
    cache[raw] = result;
    return result;
  };
}

function renderTop10Clientes6h(episodes){
  const crit = episodes.filter(e=>e.hours>=6);
  const counts = {};
  crit.forEach(e=> {
    const code = extractClientCode(e.cliente);
    if(isSciCode(code)) return; // código ilegible por notación científica, se excluye
    counts[code] = (counts[code]||0)+1;
  });
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([label,value])=>({label,value}));
  renderBarChart('top10Clientes6h', top, top[0]?.value||0);
}

function renderAlertMatrix(rows){
  const alertRows = rows.filter(r=>ALERT_EVENTS.includes(r.evento));
  // per device per day count
  const dayCount = {}; // key: distrito|device|dateStr -> count
  alertRows.forEach(r=>{
    const dateStr = toDateInputVal(r.fecha);
    const key = r.locacion+'|'+r.device+'|'+dateStr;
    dayCount[key] = (dayCount[key]||0)+1;
  });
  const buckets = ['1','2','3','4','5','>5'];
  const matrix = {}; // distrito -> bucket -> count
  for(const key in dayCount){
    const [distrito] = key.split('|');
    const c = dayCount[key];
    const b = c>=6?'>5':String(c);
    matrix[distrito] = matrix[distrito] || Object.fromEntries(buckets.map(x=>[x,0]));
    matrix[distrito][b]++;
  }
  const distritos = Object.keys(matrix).sort((a,b)=> {
    const totalA = buckets.reduce((s,bk)=>s+matrix[a][bk],0);
    const totalB = buckets.reduce((s,bk)=>s+matrix[b][bk],0);
    return totalB-totalA;
  }).slice(0,12);

  const totals = Object.fromEntries(buckets.map(b=>[b,0]));
  distritos.forEach(d=> buckets.forEach(b=> totals[b]+=matrix[d][b]));
  const grandTotal = buckets.reduce((s,b)=>s+totals[b],0);
  const maxCell = Math.max(1, ...distritos.flatMap(d=>buckets.map(b=>matrix[d][b])));

  let html = `<thead><tr><th>Distrito</th>${buckets.map(b=>`<th>${b} alerta${b==='1'?'':'s'}</th>`).join('')}<th>Total</th></tr></thead><tbody>`;
  distritos.forEach(d=>{
    const rowTotal = buckets.reduce((s,b)=>s+matrix[d][b],0);
    html += `<tr><td>${d}</td>${buckets.map(b=>{
      const v = matrix[d][b];
      return heatDot(v, maxCell);
    }).join('')}<td style="font-weight:600;">${rowTotal}</td></tr>`;
  });
  html += `<tr class="total-row"><td>Total</td>${buckets.map(b=>`<td>${totals[b]}</td>`).join('')}<td>${grandTotal}</td></tr>`;
  html += '</tbody>';
  document.getElementById('tblAlertMatrix').innerHTML = html || '<tbody><tr><td>Sin datos</td></tr></tbody>';
}

function renderHorarioMatrix(episodes){
  const matrix = {}; // horario -> dur -> count
  HORARIO_BUCKETS.forEach(h=> matrix[h] = Object.fromEntries(DUR_BUCKETS.map(d=>[d,0])));
  episodes.forEach(e=>{
    matrix[horarioBucket(e.start)][durationBucket(e.hours)]++;
  });
  const totals = Object.fromEntries(DUR_BUCKETS.map(d=>[d,0]));
  HORARIO_BUCKETS.forEach(h=> DUR_BUCKETS.forEach(d=> totals[d]+=matrix[h][d]));
  const grandTotal = DUR_BUCKETS.reduce((s,d)=>s+totals[d],0);
  const maxCell = Math.max(1, ...HORARIO_BUCKETS.flatMap(h=>DUR_BUCKETS.map(d=>matrix[h][d])));

  let html = `<thead><tr><th>Horario</th>${DUR_BUCKETS.map(d=>`<th>${d}</th>`).join('')}<th>Total</th></tr></thead><tbody>`;
  HORARIO_BUCKETS.forEach(h=>{
    const rowTotal = DUR_BUCKETS.reduce((s,d)=>s+matrix[h][d],0);
    html += `<tr><td>${h}</td>${DUR_BUCKETS.map(d=>{
      const v = matrix[h][d];
      return heatDot(v, maxCell);
    }).join('')}<td style="font-weight:600;">${rowTotal}</td></tr>`;
  });
  html += `<tr class="total-row"><td>Total</td>${DUR_BUCKETS.map(d=>`<td>${totals[d]}</td>`).join('')}<td>${grandTotal}</td></tr>`;
  html += '</tbody>';
  document.getElementById('tblHorarioMatrix').innerHTML = html;
}

function renderDistritoMatrix(episodes){
  const matrix = {};
  episodes.forEach(e=>{
    matrix[e.locacion] = matrix[e.locacion] || Object.fromEntries(DUR_BUCKETS.map(d=>[d,0]));
    matrix[e.locacion][durationBucket(e.hours)]++;
  });
  const distritos = Object.keys(matrix).sort((a,b)=>{
    const totalA = DUR_BUCKETS.reduce((s,d)=>s+matrix[a][d],0);
    const totalB = DUR_BUCKETS.reduce((s,d)=>s+matrix[b][d],0);
    return totalB-totalA;
  }).slice(0,12);
  const totals = Object.fromEntries(DUR_BUCKETS.map(d=>[d,0]));
  distritos.forEach(dt=> DUR_BUCKETS.forEach(d=> totals[d]+=matrix[dt][d]));
  const grandTotal = DUR_BUCKETS.reduce((s,d)=>s+totals[d],0);
  const maxCell = Math.max(1, ...distritos.flatMap(dt=>DUR_BUCKETS.map(d=>matrix[dt][d])));

  let html = `<thead><tr><th>Distrito</th>${DUR_BUCKETS.map(d=>`<th>${d}</th>`).join('')}<th>Total</th></tr></thead><tbody>`;
  distritos.forEach(dt=>{
    const rowTotal = DUR_BUCKETS.reduce((s,d)=>s+matrix[dt][d],0);
    html += `<tr><td>${dt}</td>${DUR_BUCKETS.map(d=>{
      const v = matrix[dt][d];
      return heatDot(v, maxCell);
    }).join('')}<td style="font-weight:600;">${rowTotal}</td></tr>`;
  });
  html += `<tr class="total-row"><td>Total</td>${DUR_BUCKETS.map(d=>`<td>${totals[d]}</td>`).join('')}<td>${grandTotal}</td></tr>`;
  html += '</tbody>';
  document.getElementById('tblDistritoMatrix').innerHTML = html;
}

// ---------------- Tabs (Resumen general / Detalle por placa, DENTRO del módulo Desconexiones) ----------------
// OJO: usamos '.tabbar:not(.module-tabbar) .tab-btn' para no capturar también los
// botones de arriba (Censo / Desconexiones), que usan la misma clase .tab-btn pero
// tienen su propio manejador más abajo (ver "Tabbar de módulo"). Antes esto colisionaba:
// este listener genérico también se disparaba en los botones de módulo y fallaba
// buscando un elemento "tab-undefined" que no existe.
document.querySelectorAll('.tabbar:not(.module-tabbar) .tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    const scope = btn.closest('.tabbar');
    scope.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    document.querySelectorAll('.tab-panel').forEach(p=>p.style.display='none');
    document.getElementById('tab-'+btn.dataset.tab).style.display='block';
  });
});

// ---------------- Detalle por placa ----------------
function fmtDuracion(h){
  if(h === null || h === undefined || isNaN(h)) return '–';
  const totalMin = Math.round(h*60);
  const hh = Math.floor(totalMin/60);
  const mm = totalMin%60;
  if(hh === 0) return `${mm} min`;
  if(mm === 0) return `${hh} h`;
  return `${hh} h ${mm} min`;
}

function initPlacaUI(){
  const devices = uniqueSorted([...new Set(RAW.map(r=>r.device))]);
  const dl = document.getElementById('placaList');
  dl.innerHTML = devices.map(d=>`<option value="${d}">`).join('');
  const input = document.getElementById('placaSearch');
  input.value = '';
  document.getElementById('placaEmpty').style.display = 'flex';
  document.getElementById('placaContent').style.display = 'none';
  input.oninput = () => {
    const val = input.value.trim();
    if(devices.includes(val)){
      renderPlacaDetalle(val);
    } else {
      document.getElementById('placaEmpty').style.display = 'flex';
      document.getElementById('placaContent').style.display = 'none';
    }
  };
}

let currentPlacaDevice = null;

function clientNamePart(clienteText, code){
  return clienteText.trim().slice(code.length).replace(/^\S*\s*/,'').trim() || clienteText.trim().replace(/^\S+\s*/,'').trim();
}

function renderPlacaDetalle(device){
  currentPlacaDevice = device;
  document.getElementById('placaEmpty').style.display = 'none';
  document.getElementById('placaContent').style.display = 'block';

  const deviceRows = RAW.filter(r=>r.device === device); // ya viene ordenado por fecha global
  const resolveCode = buildCodeResolver(deviceRows);
  const periods = [];
  deviceRows.forEach(r=>{
    const rawCode0 = extractClientCode(r.cliente);
    const rawCode = resolveCode(rawCode0);
    const rawName = isSciCode(rawCode0) ? '' : clientNamePart(r.cliente, rawCode0);
    const last = periods[periods.length-1];

    let sameClient = false;
    let code = rawCode;
    if(last){
      if(isSciCode(rawCode)){
        sameClient = true; code = last.code; // no se pudo resolver, se asume continuidad
      } else if(last.code === rawCode){
        if(!last.name || !rawName || last.name === rawName){
          sameClient = true; code = last.code;
        }
      }
    }

    if(sameClient){
      last.hasta = r.fecha;
      if(rawName && r.cliente.length > last.cliente.length){ last.cliente = r.cliente; last.name = rawName; }
    } else {
      periods.push({code, name:rawName, cliente:r.cliente, locacion:r.locacion, desde:r.fecha, hasta:r.fecha});
    }
  });

  const groups = {}; // key -> {label, periods:[]}
  periods.forEach(p=>{
    const key = p.code + '|' + p.name;
    if(!groups[key]) groups[key] = {label:p.cliente, periods:[]};
    if(p.cliente.length > groups[key].label.length) groups[key].label = p.cliente;
    groups[key].periods.push(p);
  });
  const groupKeys = Object.keys(groups);

  document.getElementById('pkPlaca').textContent = device;
  if(groupKeys.length > 1){
    document.getElementById('pkCliente').innerHTML = `<span style="color:var(--amber); font-weight:600;">⚠ ${groupKeys.length} clientes distintos</span> · ver tabla abajo`;
  } else {
    document.getElementById('pkCliente').textContent = periods[0] ? `${periods[0].cliente} · ${periods[0].locacion}` : 'Sin registros';
  }

  const allEps = EPISODES.filter(e=>e.device === device);
  let clHtml = `<thead><tr><th>Cliente</th><th>Distrito</th><th>Desde</th><th>Hasta</th><th>Incidentes en ese periodo</th></tr></thead><tbody>`;
  periods.slice().reverse().forEach(p=>{
    const incidentesPeriodo = allEps.filter(e=> e.start >= p.desde && e.start <= p.hasta).length;
    clHtml += `<tr>
      <td>${p.cliente}</td>
      <td>${p.locacion}</td>
      <td>${fmtDate(p.desde)}</td>
      <td>${fmtDate(p.hasta)}</td>
      <td>${incidentesPeriodo}</td>
    </tr>`;
  });
  clHtml += '</tbody>';
  document.getElementById('tblPlacaClientes').innerHTML = clHtml;
  document.getElementById('clientesPanelWrap').style.display = groupKeys.length > 1 ? 'block' : 'none';

  const filterWrap = document.getElementById('placaClienteFilterWrap');
  const select = document.getElementById('placaClienteSelect');
  if(groupKeys.length > 1){
    filterWrap.style.display = 'flex';
    select.innerHTML = `<option value="__ALL__">Todos los clientes (${groupKeys.length})</option>` +
      groupKeys.map(key=>`<option value="${key.replace(/"/g,'&quot;')}">${groups[key].label}</option>`).join('');
    select.value = '__ALL__';
    select.onchange = () => renderPlacaHistorial(device, select.value, groups);
  } else {
    filterWrap.style.display = 'none';
  }

  renderPlacaHistorial(device, '__ALL__', groups);
}

function renderPlacaHistorial(device, clienteFilter, groups){
  let eps = EPISODES.filter(e=>e.device === device);
  if(clienteFilter && clienteFilter !== '__ALL__' && groups && groups[clienteFilter]){
    const ranges = groups[clienteFilter].periods;
    eps = eps.filter(e => ranges.some(p => e.start >= p.desde && e.start <= p.hasta));
  }
  eps = eps.sort((a,b)=>b.start-a.start);

  document.getElementById('pkTotal').textContent = eps.length.toLocaleString('es-PE');
  const totalHoras = eps.reduce((s,e)=>s+e.hours,0);
  document.getElementById('pkHoras').textContent = totalHoras.toFixed(1)+'h';
  document.getElementById('pkCriticos').textContent = eps.filter(e=>e.hours>=6).length.toLocaleString('es-PE');
  if(eps.length){
    const worst = eps.reduce((a,b)=> b.hours>a.hours? b:a, eps[0]);
    document.getElementById('pkMax').textContent = fmtDuracion(worst.hours);
    document.getElementById('pkMaxDate').textContent = fmtDate(worst.start);
  } else {
    document.getElementById('pkMax').textContent = '–';
    document.getElementById('pkMaxDate').textContent = '–';
  }
  const filtroTxt = (clienteFilter && clienteFilter !== '__ALL__') ? ' (filtrado por cliente)' : '';
  document.getElementById('pkBadge').textContent = `${eps.length} corte${eps.length===1?'':'s'}${filtroTxt}`;

  let html = `<thead><tr><th>Fecha</th><th>Hora inicio</th><th>Hora fin</th><th>Duración</th><th>Cliente</th><th>Se originó por</th><th>Distrito</th><th>Estado</th></tr></thead><tbody>`;
  if(eps.length === 0){
    html += `<tr><td colspan="8" style="text-align:center; color:var(--muted2); padding:18px;">No hay incidentes para este filtro.</td></tr>`;
  }
  eps.forEach(e=>{
    const durColor = e.hours>=6 ? 'var(--red)' : (e.hours>=1 ? 'var(--amber)' : 'var(--muted)');
    html += `<tr>
      <td>${fmtDate(e.start)}</td>
      <td>${pad(e.start.getHours())}:${pad(e.start.getMinutes())}</td>
      <td>${e.end ? pad(e.end.getHours())+':'+pad(e.end.getMinutes()) : '—'}</td>
      <td style="color:${durColor}; font-weight:600;">${fmtDuracion(e.hours)}</td>
      <td>${e.cliente}</td>
      <td>${e.startEvento||'–'}</td>
      <td>${e.locacion}</td>
      <td>${e.ongoing ? '<span style="color:var(--red);">En curso</span>' : 'Cerrado'}</td>
    </tr>`;
  });
  html += '</tbody>';
  document.getElementById('tblPlacaHistorial').innerHTML = html;
}

// =====================================================================
// MÓDULO: CENSO DE ACTIVOS
// =====================================================================
let RAW_CENSO = [];
let fechaArchivoCenso = null; // fecha real del archivo (Last-Modified de GitHub)
let censoMap = null;
let censoMarkersLayer = null;

function pickLargestSheet(wb){
  let best = wb.SheetNames[0], bestRows = -1;
  wb.SheetNames.forEach(name=>{
    const ref = wb.Sheets[name]['!ref'];
    if(!ref) return;
    const range = XLSX.utils.decode_range(ref);
    const rows = range.e.r - range.s.r;
    if(rows > bestRows){ bestRows = rows; best = name; }
  });
  return best;
}

async function fetchAndLoadCenso(){
  const mapEl = document.getElementById('censoMap');
  mapEl.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted);font-family:\'IBM Plex Mono\',monospace;font-size:13px;gap:8px;"><div class="spinner"></div>Cargando censo de activos…</div>';
  setSourceStatus('Censo', 'loading', 'cargando…');
  try{
    const sep = DATA_SOURCE_URL_CENSO.includes('?') ? '&' : '?';
    const res = await fetch(DATA_SOURCE_URL_CENSO + sep + 'cachebust=' + Date.now());
    if(!res.ok) throw new Error('No se pudo descargar el archivo de censo (HTTP ' + res.status + ')');
    fechaArchivoCenso = getFechaArchivo(res);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array', cellDates:true});
    // Los encabezados reales de este archivo están en la FILA 2 (no en la 1,
    // que trae títulos combinados). Por eso usamos range:1 (0-indexado).
    const sheetName = wb.SheetNames.includes('Hoja1') ? 'Hoja1' : pickLargestSheet(wb);
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, {range:1, defval:'', cellDates:true});
    normalizeCenso(data);
    if(RAW_CENSO.length === 0){
      throw new Error('El archivo no tiene filas reconocibles. Revisa que la hoja "Hoja1" tenga los encabezados esperados (Placa Física, Vendedor, Canal, STATUS, X, Y...).');
    }
    initCensoFiltersUI();
    applyCensoFiltersAndRender();
    setSourceStatus('Censo', 'ok', fechaArchivoCenso ? `archivo del ${fmtDateCorta(fechaArchivoCenso)}` : `cargado · ${RAW_CENSO.length.toLocaleString('es-PE')} activos`);
  } catch(err){
    console.error('Error cargando censo:', err);
    setSourceStatus('Censo', 'error', 'error al cargar');
    mapEl.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:var(--red);text-align:center;padding:20px;gap:10px;">
      <div>No se pudo cargar el archivo de censo.</div>
      <div style="font-size:12px;color:var(--muted); max-width:420px;">${err.message}</div>
      <button class="btn btn-primary" id="censoRetryBtn">↻ Reintentar</button>
    </div>`;
    document.getElementById('censoRetryBtn')?.addEventListener('click', fetchAndLoadCenso);
  }
}

function normalizeCenso(records){
  RAW_CENSO = records.map(r=>{
    const razonSocial = (r['Razon Social']||'').toString().trim();
    const tdRaf = (r['TD RAF']||'').toString().trim();
    const esPatio = tdRaf.toUpperCase()==='PATIOS' || !razonSocial;
    const status = (r['STATUS']||'Sin status').toString().trim() || 'Sin status';
    // "OK" cubre tanto "CENSO OK" como "TRASPASO DISMAC (CENSO OK)"
    const censado = /OK/i.test(status);
    const lat = parseFloat(r['X']);
    const lng = parseFloat(r['Y']);
    const tieneCoord = isFinite(lat) && isFinite(lng) && lat !== 0 && lng !== 0;
    let fechaUbicacion = r['FECHA DE UBICACIÓN'];
    if(!(fechaUbicacion instanceof Date) || isNaN(fechaUbicacion?.getTime?.())) fechaUbicacion = null;
    const placa = (r['PLACA FÍSICA'] || r['PLACA SISTEMA'] || '').toString().trim();
    const tipoActivo = (r['Tipo de Activo']||'Sin tipo').toString().trim() || 'Sin tipo';
    // El Excel no trae "Triciclo" como valor de Canal (viene mezclado en Ambulatorio o vacío),
    // pero operativamente sí es un canal propio — lo separamos usando el Tipo de Activo.
    let canal = (r['Canal']||'').toString().trim() || 'Sin canal';
    if(/TRICICLO/i.test(tipoActivo)) canal = 'Triciclo';
    return {
      activoFijo: r['Activo fijo'],
      placa,
      tipoActivo,
      marca: (r['MARCA']||'').toString().trim(),
      modelo: (r['MODELO']||'').toString().trim(),
      tamano: (r['TAMAÑO']||'').toString().trim(),
      cliente: esPatio ? 'Patio Sagadis' : razonSocial,
      esPatio,
      vendedor: (r['Vendedor']||'').toString().trim() || 'Sin vendedor',
      canal,
      canalVenta: (r['Canal Venta']||'').toString().trim(),
      distrito: (r['Distrito']||'').toString().trim() || 'Sin distrito',
      status,
      censado,
      fechaUbicacion,
      lat, lng, tieneCoord
    };
  }).filter(r=>r.placa);
}

let msVendedor, msCanal, msStatus, msDistritoCenso;

function initCensoFiltersUI(){
  if(!msVendedor) msVendedor = createMultiSelect('cFVendedor', {placeholder:'Todos', onChange: applyCensoFiltersAndRender});
  if(!msCanal) msCanal = createMultiSelect('cFCanal', {placeholder:'Todos', onChange: applyCensoFiltersAndRender});
  if(!msStatus) msStatus = createMultiSelect('cFStatus', {placeholder:'Todos', onChange: applyCensoFiltersAndRender});
  if(!msDistritoCenso) msDistritoCenso = createMultiSelect('cFDistrito', {placeholder:'Todos', onChange: applyCensoFiltersAndRender});
  msVendedor.setOptions(uniqueSorted(RAW_CENSO.map(r=>r.vendedor)));
  msCanal.setOptions(uniqueSorted(RAW_CENSO.map(r=>r.canal)));
  msStatus.setOptions(uniqueSorted(RAW_CENSO.map(r=>r.status)));
  msDistritoCenso.setOptions(uniqueSorted(RAW_CENSO.map(r=>r.distrito)));

  ['cFSearch','cFOnlyPatio'].forEach(id=>{
    document.getElementById(id).addEventListener('input', applyCensoFiltersAndRender);
  });
  document.getElementById('censoClearFilters').onclick = () => {
    msVendedor.clear();
    msCanal.clear();
    msStatus.clear();
    msDistritoCenso.clear();
    document.getElementById('cFSearch').value = '';
    document.getElementById('cFOnlyPatio').checked = false;
    applyCensoFiltersAndRender();
  };
}

function getCensoFilters(){
  return {
    vendedor: msVendedor ? msVendedor.getValues() : null,
    canal: msCanal ? msCanal.getValues() : null,
    status: msStatus ? msStatus.getValues() : null,
    distrito: msDistritoCenso ? msDistritoCenso.getValues() : null,
    search: document.getElementById('cFSearch').value.trim().toLowerCase(),
    onlyPatio: document.getElementById('cFOnlyPatio').checked
  };
}

function matchesCenso(r, f){
  if(f.vendedor !== null && !f.vendedor.includes(r.vendedor)) return false;
  if(f.canal !== null && !f.canal.includes(r.canal)) return false;
  if(f.status !== null && !f.status.includes(r.status)) return false;
  if(f.distrito !== null && !f.distrito.includes(r.distrito)) return false;
  if(f.onlyPatio && !r.esPatio) return false;
  if(f.search){
    const s = f.search;
    if(!r.placa.toLowerCase().includes(s) && !r.cliente.toLowerCase().includes(s)) return false;
  }
  return true;
}

function applyCensoFiltersAndRender(){
  const f = getCensoFilters();
  const rows = RAW_CENSO.filter(r=>matchesCenso(r,f));
  renderCensoKPIs(rows);
  renderCensoAvance('censoAvanceCanal', rows, 'canal', null);
  renderCensoAvance('censoAvanceVendedor', rows, 'vendedor', null);
  renderCensoAvance('censoAvanceDistrito', rows, 'distrito', null);
  renderCensoAntiguedad(rows);
  renderCensoMap(rows);
  renderCensoPendientesTable(rows);
  renderCensoActiveFilters(f);
}

// ---------------- Activos pendientes (tabla de detalle) ----------------
function renderCensoPendientesTable(rows){
  const pend = rows.filter(r=>!r.censado).slice().sort((a,b)=>{
    const da = a.fechaUbicacion ? a.fechaUbicacion.getTime() : 0;
    const db = b.fechaUbicacion ? b.fechaUbicacion.getTime() : 0;
    return da - db; // los más antiguos (más urgentes) primero
  });
  document.getElementById('censoPendientesBadge').textContent = `${pend.length.toLocaleString('es-PE')} pendientes`;
  const MAX_ROWS = 300; // límite de filas pintadas en la tabla, por rendimiento del navegador
  const now = new Date();
  let html = '<thead><tr><th>Placa</th><th>Cliente</th><th>Canal</th><th>Vendedor</th><th>Distrito</th><th>Status</th><th>Días pendiente</th></tr></thead><tbody>';
  if(pend.length === 0){
    html += '<tr><td colspan="7" style="text-align:center; color:var(--muted2); padding:18px;">No hay activos pendientes con estos filtros 🎉</td></tr>';
  }
  pend.slice(0, MAX_ROWS).forEach(r=>{
    const dias = r.fechaUbicacion ? Math.floor((now - r.fechaUbicacion)/86400000) : null;
    const diasTxt = dias === null ? 'Sin fecha' : `${dias} días`;
    const diasColor = dias === null ? 'var(--muted2)' : dias >= 90 ? 'var(--red)' : dias >= 30 ? 'var(--amber)' : 'var(--muted)';
    html += `<tr>
      <td class="mono">${r.placa || '–'}</td>
      <td>${r.cliente}</td>
      <td>${r.canal}</td>
      <td>${r.vendedor}</td>
      <td>${r.distrito}</td>
      <td>${r.status}</td>
      <td style="color:${diasColor}; font-weight:600;">${diasTxt}</td>
    </tr>`;
  });
  html += '</tbody>';
  document.getElementById('tblCensoPendientes').innerHTML = html;
  if(pend.length > MAX_ROWS){
    document.getElementById('tblCensoPendientes').insertAdjacentHTML('afterend',
      `<div style="text-align:center; font-size:11.5px; color:var(--muted2); padding:8px;" id="censoPendientesMoreNote">
        Mostrando los ${MAX_ROWS} más antiguos de ${pend.length.toLocaleString('es-PE')} — usa los filtros o el buscador para acotar.
      </div>`);
  } else {
    document.getElementById('censoPendientesMoreNote')?.remove();
  }
}

// ---------------- Filtros activos (chips visibles y removibles) ----------------
const CENSO_FILTER_LABELS = { vendedor:'Vendedor', canal:'Canal', status:'Status', distrito:'Distrito' };
const CENSO_FILTER_MSEL = () => ({ vendedor: msVendedor, canal: msCanal, status: msStatus, distrito: msDistritoCenso });

function renderCensoActiveFilters(f){
  const wrap = document.getElementById('censoActiveFilters');
  const chips = [];
  const msels = CENSO_FILTER_MSEL();
  ['vendedor','canal','status','distrito'].forEach(field=>{
    const vals = f[field];
    if(vals === null) return; // null = "Todos", no se muestra chip
    if(vals.length === 0){
      chips.push({ text: `${CENSO_FILTER_LABELS[field]}: ninguno seleccionado`, onClear: () => msels[field].clear() });
    } else if(vals.length <= 2){
      chips.push({ text: `${CENSO_FILTER_LABELS[field]}: ${vals.join(', ')}`, onClear: () => msels[field].clear() });
    } else {
      chips.push({ text: `${CENSO_FILTER_LABELS[field]}: ${vals.length} seleccionados`, onClear: () => msels[field].clear() });
    }
  });
  if(f.onlyPatio){
    chips.push({ text: 'Solo patio Sagadis', onClear: () => { document.getElementById('cFOnlyPatio').checked = false; } });
  }
  if(f.search){
    chips.push({ text: `Búsqueda: "${f.search}"`, onClear: () => { document.getElementById('cFSearch').value = ''; } });
  }
  if(chips.length === 0){ wrap.innerHTML = ''; return; }
  wrap.innerHTML = chips.map((c,i)=>`
    <span class="filter-chip">${c.text}<span class="chip-x" data-chip="${i}">✕</span></span>
  `).join('');
  wrap.querySelectorAll('.chip-x').forEach(el=>{
    el.addEventListener('click', ()=>{
      chips[+el.dataset.chip].onClear();
      applyCensoFiltersAndRender(); // .clear() de los msel no dispara onChange por sí solo
    });
  });
}

// Clic en cualquier barra de "Avance por Canal / Vendedor / Distrito": filtra por
// ese valor y baja la vista hasta la tabla de activos pendientes correspondiente.
document.addEventListener('click', (e)=>{
  const bar = e.target.closest('.bar-clickable');
  if(!bar) return;
  const field = bar.dataset.field; // 'canal' | 'vendedor' | 'distrito'
  const label = bar.dataset.label;
  const msels = CENSO_FILTER_MSEL();
  const target = msels[field];
  if(!target) return;
  target.selectOnly(label);
  applyCensoFiltersAndRender();
  document.getElementById('censoPendientesPanel')?.scrollIntoView({ behavior:'smooth', block:'start' });
});

function fmtPct(n){ return n.toFixed(1) + '%'; } // 1 decimal siempre (ej. 99.9%, nunca redondea a 100%)

function renderCensoKPIs(rows){
  const total = rows.length;
  const censados = rows.filter(r=>r.censado).length;
  const pendientes = total - censados;
  const pct = total ? (censados/total*100) : 0;
  const pctPend = total ? (pendientes/total*100) : 0;
  const now = new Date();
  const viejos = rows.filter(r=> !r.censado && r.fechaUbicacion && (now - r.fechaUbicacion)/86400000 >= 30).length;
  const patio = rows.filter(r=>r.esPatio).length;
  document.getElementById('cKpiTotal').textContent = total.toLocaleString('es-PE');
  document.getElementById('cKpiTotalSub').textContent = `de ${RAW_CENSO.length.toLocaleString('es-PE')} totales`;
  document.getElementById('cKpiCensados').textContent = censados.toLocaleString('es-PE');
  document.getElementById('cKpiPct').textContent = `${fmtPct(pct)} de avance`;
  document.getElementById('cKpiPendientes').textContent = pendientes.toLocaleString('es-PE');
  document.getElementById('cKpiPendientesPct').textContent = `${fmtPct(pctPend)} del total`;
  document.getElementById('cKpiViejos').textContent = viejos.toLocaleString('es-PE');
  document.getElementById('cKpiPatio').textContent = patio.toLocaleString('es-PE');
}

function renderCensoAvance(containerId, rows, field, topN){
  const groups = {};
  rows.forEach(r=>{
    const k = r[field];
    if(!groups[k]) groups[k] = {total:0, censados:0, pendientesRows:[]};
    groups[k].total++;
    if(r.censado) groups[k].censados++; else groups[k].pendientesRows.push(r);
  });
  const items = Object.entries(groups)
    .map(([label,g])=>({ label, pct: g.total ? (g.censados/g.total*100) : 0, total:g.total, censados:g.censados, pendientes:g.total-g.censados }))
    .sort((a,b)=> b.total - a.total);
  const limited = (topN ? items.slice(0, topN) : items).sort((a,b)=> b.pct - a.pct);
  const container = document.getElementById(containerId);
  if(items.length === 0){
    container.innerHTML = '<div style="color:var(--muted2); font-size:12.5px;">Sin datos en el rango filtrado.</div>';
    return;
  }
  container.innerHTML = limited.map(it=>{
    const color = it.pct >= 95 ? '#2e9e4f' : it.pct >= 80 ? 'var(--gold)' : 'var(--red)';
    const tip = `${it.label}: ${it.pendientes.toLocaleString('es-PE')} pendientes de ${it.total.toLocaleString('es-PE')} activos — clic para ver el detalle`;
    return `
    <div class="bar-wrap bar-clickable" data-field="${field}" data-label="${it.label.replace(/"/g,'&quot;')}" title="${tip}">
      <div class="bar-label">${it.label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${it.pct}%; background:${color};"></div></div>
      <div class="bar-val">${fmtPct(it.pct)}</div>
    </div>`;
  }).join('');
}

function renderCensoAntiguedad(rows){
  const now = new Date();
  const buckets = {'Menos de 30 días':0, 'Entre 30 y 60 días':0, 'Entre 60 y 90 días':0, 'Mayor a 90 días':0, 'Sin fecha de ubicación':0};
  rows.filter(r=>!r.censado).forEach(r=>{
    if(!r.fechaUbicacion){ buckets['Sin fecha de ubicación']++; return; }
    const dias = Math.floor((now - r.fechaUbicacion)/86400000);
    if(dias < 30) buckets['Menos de 30 días']++;
    else if(dias < 60) buckets['Entre 30 y 60 días']++;
    else if(dias < 90) buckets['Entre 60 y 90 días']++;
    else buckets['Mayor a 90 días']++;
  });
  const total = Object.values(buckets).reduce((a,b)=>a+b,0);
  let html = '<thead><tr><th>Días desde ubicación</th><th>N° Placas</th></tr></thead><tbody>';
  Object.entries(buckets).forEach(([k,v])=>{
    if(v === 0) return;
    html += `<tr><td>${k}</td><td style="font-weight:600;">${v.toLocaleString('es-PE')}</td></tr>`;
  });
  html += `<tr class="total-row"><td>Total pendientes</td><td>${total.toLocaleString('es-PE')}</td></tr></tbody>`;
  document.getElementById('tblCensoAntiguedad').innerHTML = total ? html : '<tbody><tr><td>No hay activos pendientes en el rango filtrado 🎉</td></tr></tbody>';
}

function makeDotIcon(color){
  return L.divIcon({
    className: '',
    html: `<div style="width:13px;height:13px;border-radius:50%;background:${color};border:2px solid #fff;box-shadow:0 0 0 1px rgba(0,0,0,.25);"></div>`,
    iconSize: [13,13],
    iconAnchor: [6.5,6.5]
  });
}

function renderCensoMap(rows){
  const withCoords = rows.filter(r=>r.tieneCoord);
  document.getElementById('censoMapBadge').textContent = `${withCoords.length.toLocaleString('es-PE')} activos con ubicación en el mapa`;

  if(!censoMap){
    document.getElementById('censoMap').innerHTML = '';
    censoMap = L.map('censoMap', { zoomControl: true, attributionControl: true });
    // Nota: CARTO (basemaps.cartocdn.com) empezó a exigir una API key incluso para
    // uso gratuito — se cambió a OpenStreetMap estándar, que no requiere ninguna
    // clave ni registro.
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
      maxZoom: 19, subdomains: 'abc'
    }).addTo(censoMap);

    censoMarkersLayer = L.markerClusterGroup({
      maxClusterRadius: 50,
      disableClusteringAtZoom: 17,
      iconCreateFunction: function(cluster){
        const children = cluster.getAllChildMarkers();
        const total = children.length;
        const ok = children.filter(m=>m.options.censado).length;
        const pct = total ? Math.round(ok/total*100) : 0;
        const color = pct >= 90 ? '#2e9e4f' : pct >= 70 ? '#c67c0a' : '#c0392b';
        const size = total < 20 ? 34 : total < 100 ? 42 : 52;
        return L.divIcon({
          html: `<div style="width:${size}px;height:${size}px;border-radius:50%;background:${color};color:#fff;display:flex;align-items:center;justify-content:center;font-family:'IBM Plex Mono',monospace;font-weight:700;font-size:${size>44?13:11.5}px;border:3px solid rgba(255,255,255,.85);box-shadow:0 2px 8px rgba(16,24,48,.35);">${total}</div>`,
          className: '',
          iconSize: [size, size]
        });
      }
    });
    censoMap.addLayer(censoMarkersLayer);

    const legend = L.control({position:'bottomleft'});
    legend.onAdd = function(){
      const div = L.DomUtil.create('div', 'censo-map-legend');
      div.innerHTML = `
        <div><span class="dot" style="background:#2e9e4f;"></span>Censado</div>
        <div><span class="dot" style="background:#c0392b;"></span>Pendiente</div>
        <div style="margin-top:4px; color:var(--muted2); font-size:10.5px;">Círculos: grupos de activos · el color indica su % de avance</div>
      `;
      return div;
    };
    legend.addTo(censoMap);
  } else {
    censoMarkersLayer.clearLayers();
  }

  const iconOk = makeDotIcon('#2e9e4f');
  const iconPend = makeDotIcon('#c0392b');
  const markers = withCoords.map(r=>{
    const marker = L.marker([r.lat, r.lng], { icon: r.censado ? iconOk : iconPend, censado: r.censado });
    marker.bindPopup(`
      <div class="censo-popup">
        <b>${r.placa || 'Sin placa'}</b> · ${r.tipoActivo}<br>
        ${r.cliente}<br>
        Canal: ${r.canal} · Vendedor: ${r.vendedor}<br>
        Distrito: ${r.distrito}<br>
        Status: ${r.censado ? '<span class="tag-ok">'+r.status+'</span>' : '<span class="tag-pend">'+r.status+'</span>'}
      </div>
    `);
    return marker;
  });
  censoMarkersLayer.addLayers(markers);

  if(!censoMap._encuadreInicialHecho){
    if(markers.length){
      censoMap.fitBounds(censoMarkersLayer.getBounds(), { padding:[20,20], maxZoom:14 });
    } else {
      censoMap.setView([-12.05, -77.03], 11); // vista por defecto: Lima
    }
    censoMap._encuadreInicialHecho = true;
  }
  setTimeout(()=>censoMap.invalidateSize(), 60);
}

// ---------------- Tabbar de módulo (Desconexiones vs Censo) ----------------
document.querySelectorAll('.module-tabbar .tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.module-tabbar .tab-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    const mod = btn.dataset.module;
    document.getElementById('module-desconexiones').style.display = mod === 'desconexiones' ? 'block' : 'none';
    document.getElementById('module-censo').style.display = mod === 'censo' ? 'block' : 'none';
    if(mod === 'censo'){
      if(RAW_CENSO.length === 0){
        fetchAndLoadCenso();
      } else if(censoMap){
        setTimeout(()=>censoMap.invalidateSize(), 60);
      }
    }
    if(mod === 'desconexiones' && RAW.length === 0){
      fetchAndLoadData(false);
    }
  });
});

// ---------------- Autenticación (Supabase) ----------------
if(typeof window.supabase === 'undefined'){
  document.getElementById('loginError').textContent = 'No se pudo cargar la librería de autenticación (Supabase). Revisa tu conexión a internet y recarga la página.';
  document.getElementById('loginError').style.display = 'block';
  console.error('window.supabase no está definido — el script de supabase-js no cargó.');
}

const supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const loginScreen = document.getElementById('loginScreen');
const appWrap = document.getElementById('appWrap');
const loginForm = document.getElementById('loginForm');
const loginError = document.getElementById('loginError');
const loginSubmitBtn = document.getElementById('loginSubmitBtn');
const logoutBtn = document.getElementById('logoutBtn');
const userEmailLabel = document.getElementById('userEmailLabel');

function showApp(session){
  loginScreen.style.display = 'none';
  appWrap.style.display = 'block';
  userEmailLabel.textContent = session?.user?.email || '';
  if(RAW_CENSO.length === 0){
    fetchAndLoadCenso();
  }
}

function showLogin(){
  appWrap.style.display = 'none';
  loginScreen.style.display = 'flex';
  loginError.style.display = 'none';
  loginForm.reset();
}

supabaseClient.auth.getSession().then(({data:{session}})=>{
  if(session){ showApp(session); } else { showLogin(); }
}).catch(err=>{
  console.error('Error revisando sesión existente:', err);
  showLogin();
});

supabaseClient.auth.onAuthStateChange((event, session) => {
  if(event === 'SIGNED_IN' && session){ showApp(session); }
  if(event === 'SIGNED_OUT'){ showLogin(); }
});

loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  loginError.style.display = 'none';
  loginSubmitBtn.disabled = true;
  document.getElementById('loginSpinner').style.display = 'inline-block';
  document.getElementById('loginBtnText').textContent = 'Ingresando…';
  const email = document.getElementById('loginEmail').value.trim();
  const password = document.getElementById('loginPassword').value;
  try{
    const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
    if(error){
      console.error('Error de login de Supabase:', error);
      loginError.textContent = error.message.includes('Invalid login credentials')
        ? 'Correo o contraseña incorrectos.'
        : 'No se pudo iniciar sesión: ' + error.message;
      loginError.style.display = 'block';
    }
  } catch(err){
    console.error('Excepción inesperada al iniciar sesión:', err);
    loginError.textContent = 'Error de conexión al intentar iniciar sesión: ' + err.message + '. Abre la consola del navegador (F12) para más detalle.';
    loginError.style.display = 'block';
  } finally {
    loginSubmitBtn.disabled = false;
    document.getElementById('loginSpinner').style.display = 'none';
    document.getElementById('loginBtnText').textContent = 'Ingresar';
  }
});

logoutBtn.addEventListener('click', async () => {
  try{
    await supabaseClient.auth.signOut();
  } catch(err){
    console.error('Error al cerrar sesión:', err);
  }
});
