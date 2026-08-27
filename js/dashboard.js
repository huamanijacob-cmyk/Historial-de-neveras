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
const dashboardEl = document.getElementById('dashboard');
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
async function fetchAndLoadData(isRetry){
  loader.style.display = 'flex';
  document.getElementById('loaderText').textContent = isRetry ? 'Actualizando datos…' : 'Cargando datos…';
  emptyLoadBtn.style.display = 'none';
  document.getElementById('emptyTitle').textContent = isRetry ? 'Actualizando datos…' : 'Cargando datos…';
  document.getElementById('emptyText').textContent = 'Obteniendo el archivo más reciente desde el repositorio.';

  try{
    // cachebust evita que el navegador muestre una copia vieja en caché
    const sep = DATA_SOURCE_URL.includes('?') ? '&' : '?';
    const res = await fetch(DATA_SOURCE_URL + sep + 'cachebust=' + Date.now());
    if(!res.ok) throw new Error('No se pudo descargar el archivo (HTTP ' + res.status + ')');

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
    document.getElementById('fileMeta').textContent = `${loadedFileName} · ${RAW.length.toLocaleString('es-PE')} filas`;
    initFiltersUI();
    initPlacaUI();
    applyFiltersAndRender();
    emptyState.style.display = 'none';
    dashboardEl.style.display = 'block';
    document.getElementById('liveDot').style.display = 'inline-block';
    loader.style.display = 'none';
  }, 30);
}

// ---------------- Filters ----------------
function uniqueSorted(arr){ return [...new Set(arr)].sort((a,b)=>a.localeCompare(b,'es')); }

function initFiltersUI(){
  const distribuidores = uniqueSorted(RAW.map(r=>r.distribuidor));
  const distritos = uniqueSorted(RAW.map(r=>r.locacion));
  const eventos = uniqueSorted(RAW.map(r=>r.evento));
  const maxDate = RAW[RAW.length-1].fecha;
  const minDate = RAW[0].fecha;

  fillSelect('fDistribuidor', distribuidores, 'Todos');
  fillSelect('fDistrito', distritos, 'Todos');
  fillSelect('fEvento', eventos, 'Todos');
  document.getElementById('fDesde').value = toDateInputVal(minDate);
  document.getElementById('fHasta').value = toDateInputVal(maxDate);
  document.getElementById('fDesde').min = toDateInputVal(minDate);
  document.getElementById('fHasta').max = toDateInputVal(maxDate);

  ['fDistribuidor','fDistrito','fEvento','fDesde','fHasta','fSearch','fOnlyOfficial'].forEach(id=>{
    document.getElementById(id).addEventListener('input', applyFiltersAndRender);
  });
  document.getElementById('clearFilters').onclick = () => {
    fillSelect('fDistribuidor', distribuidores, 'Todos');
    fillSelect('fDistrito', distritos, 'Todos');
    fillSelect('fEvento', eventos, 'Todos');
    document.getElementById('fDesde').value = toDateInputVal(minDate);
    document.getElementById('fHasta').value = toDateInputVal(maxDate);
    document.getElementById('fSearch').value = '';
    document.getElementById('fOnlyOfficial').checked = false;
    applyFiltersAndRender();
  };
}

function fillSelect(id, values, allLabel){
  const sel = document.getElementById(id);
  sel.innerHTML = `<option value="">${allLabel}</option>` + values.map(v=>`<option value="${v.replace(/"/g,'&quot;')}">${v}</option>`).join('');
}

function getFilters(){
  return {
    distribuidor: document.getElementById('fDistribuidor').value,
    distrito: document.getElementById('fDistrito').value,
    evento: document.getElementById('fEvento').value,
    desde: document.getElementById('fDesde').value ? new Date(document.getElementById('fDesde').value+'T00:00:00') : null,
    hasta: document.getElementById('fHasta').value ? new Date(document.getElementById('fHasta').value+'T23:59:59') : null,
    search: document.getElementById('fSearch').value.trim().toLowerCase(),
    onlyOfficial: document.getElementById('fOnlyOfficial').checked
  };
}

function matchesRow(r, f){
  if(f.distribuidor && r.distribuidor !== f.distribuidor) return false;
  if(f.distrito && r.locacion !== f.distrito) return false;
  if(f.evento && r.evento !== f.evento) return false;
  if(f.desde && r.fecha < f.desde) return false;
  if(f.hasta && r.fecha > f.hasta) return false;
  if(f.search){
    const s = f.search;
    if(!r.cliente.toLowerCase().includes(s) && !r.device.toLowerCase().includes(s)) return false;
  }
  return true;
}
function matchesEpisode(ep, f){
  if(f.distribuidor && ep.distribuidor !== f.distribuidor) return false;
  if(f.distrito && ep.locacion !== f.distrito) return false;
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
  // azul marino (bajo) -> dorado (medio) -> rojo (alto)
  let bg;
  if(t < 0.5){
    const p = t/0.5;
    bg = mix('#dbe4f7','#f2c94c',p);
  } else {
    const p = (t-0.5)/0.5;
    bg = mix('#f2c94c','#c0392b',p);
  }
  const fg = t > 0.62 ? '#ffffff' : '#1b2436';
  return {bg, fg};
}

function mix(c1,c2,p){
  const a = hexToRgb(c1), b = hexToRgb(c2);
  const r = Math.round(a[0]+(b[0]-a[0])*p);
  const g = Math.round(a[1]+(b[1]-a[1])*p);
  const bl = Math.round(a[2]+(b[2]-a[2])*p);
  return `rgb(${r},${g},${bl})`;
}
function hexToRgb(h){ const n = parseInt(h.slice(1),16); return [(n>>16)&255,(n>>8)&255,n&255]; }

function renderBarList(container, items, maxVal){
  container.innerHTML = items.map(it=>`
    <div class="bar-wrap">
      <div class="bar-label" title="${it.label}">${it.label}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${maxVal? (it.value/maxVal*100):0}%"></div></div>
      <div class="bar-val">${it.value}</div>
    </div>`).join('') || '<div style="color:var(--muted2); font-size:12.5px;">Sin datos en el rango filtrado.</div>';
}

function renderTop10Activos24h(rows, maxDateAll){
  const from = new Date(maxDateAll.getTime() - 24*3600000);
  const alertRows = rows.filter(r=>ALERT_EVENTS.includes(r.evento) && r.fecha >= from && r.fecha <= maxDateAll);
  const counts = {};
  alertRows.forEach(r=> counts[r.device] = (counts[r.device]||0)+1);
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([label,value])=>({label,value}));
  document.getElementById('badge24h').textContent = `${fmtDateTime(from)} → ${fmtDateTime(maxDateAll)}`;
  renderBarList(document.getElementById('top10Activos'), top, top[0]?.value||0);
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
  renderBarList(document.getElementById('top10Clientes6h'), top, top[0]?.value||0);
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
      const hc = heatColor(v,maxCell); return `<td><span class="heat-cell" style="background:${hc.bg}; color:${hc.fg};">${v||''}</span></td>`;
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
      const hc = heatColor(v,maxCell); return `<td><span class="heat-cell" style="background:${hc.bg}; color:${hc.fg};">${v||''}</span></td>`;
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
      const hc = heatColor(v,maxCell); return `<td><span class="heat-cell" style="background:${hc.bg}; color:${hc.fg};">${v||''}</span></td>`;
    }).join('')}<td style="font-weight:600;">${rowTotal}</td></tr>`;
  });
  html += `<tr class="total-row"><td>Total</td>${DUR_BUCKETS.map(d=>`<td>${totals[d]}</td>`).join('')}<td>${grandTotal}</td></tr>`;
  html += '</tbody>';
  document.getElementById('tblDistritoMatrix').innerHTML = html;
}

// ---------------- Tabs ----------------
document.querySelectorAll('.tab-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
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

  // Reconstruye periodos de cliente en orden cronológico (todas las filas del activo,
  // no solo los incidentes). Dos filas consecutivas se consideran DEL MISMO cliente si:
  //  - su código (ya resuelto) coincide, Y
  //  - el nombre coincide, o a una de las dos le falta el nombre (fila con solo código).
  // Si el código coincide pero AMBAS filas traen nombre y son distintos (ej. mismo
  // código de cuenta pero registrado alguna vez a nombre de otra persona — pasa con
  // cuentas familiares/negocios), se trata como un cliente distinto de verdad.
  // Si el código viene dañado en notación científica (típico de Excel con números
  // largos, ej. "2.203E+13"), se intenta resolver contra los códigos completos que
  // esa misma placa mostró en otras filas antes de decidir.
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

  // Agrupa los periodos por identidad real de cliente (código+nombre), por si el mismo
  // cliente reaparece más adelante después de un tramo con otro nombre.
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

  // Tabla de clientes asociados (siempre completa, sin filtrar)
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

  // Selector de cliente para filtrar el historial — solo visible si hay más de uno real
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
    // se filtra por los rangos de fecha del grupo elegido (no comparando texto por
    // incidente, para que también cubra los incidentes cuya fila de origen no traía
    // el nombre del cliente)
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

// ---------------- Autenticación (Supabase) ----------------
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
  // Solo carga los datos la primera vez (si RAW ya tiene filas, no recarga sola al volver de otra pestaña)
  if(RAW.length === 0){
    fetchAndLoadData(false);
  }
}

function showLogin(){
  appWrap.style.display = 'none';
  loginScreen.style.display = 'flex';
  loginError.style.display = 'none';
  loginForm.reset();
}

// Revisa si ya había una sesión activa (por ejemplo, si recargaste la página)
supabaseClient.auth.getSession().then(({data:{session}})=>{
  if(session){ showApp(session); } else { showLogin(); }
});

// Reacciona a cambios de sesión (login, logout, expiración)
supabaseClient.auth.onAuthStateChange((event, session) => {
  if(event === 'SIGNED_IN' && session){ showApp(session); }
  if(event === 'SIGNED_OUT'){ showLogin(); }
});

loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  loginError.style.display = 'none';
  loginSubmitBtn.disabled = true;
  loginSubmitBtn.textContent = 'Ingresando…';
  const email = document.getElementById('loginEmail').value.trim();
  const password = document.getElementById('loginPassword').value;
  const { error } = await supabaseClient.auth.signInWithPassword({ email, password });
  loginSubmitBtn.disabled = false;
  loginSubmitBtn.textContent = 'Ingresar';
  if(error){
    loginError.textContent = error.message.includes('Invalid login credentials')
      ? 'Correo o contraseña incorrectos.'
      : 'No se pudo iniciar sesión: ' + error.message;
    loginError.style.display = 'block';
  }
  // Si fue exitoso, onAuthStateChange se encarga de mostrar el dashboard
});

logoutBtn.addEventListener('click', async () => {
  await supabaseClient.auth.signOut();
});
