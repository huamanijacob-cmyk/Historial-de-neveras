// =====================================================================
// ICONOS SVG — set propio de iconos lineales (estilo outline, 24x24),
// pensado para reemplazar los emojis que se usaban como iconografía del
// dashboard. Todos heredan color con `currentColor`, así que basta con
// envolverlos en un elemento con `color` definido (o usar --kpi-accent).
//
// Uso:
//   ICONS.bell                          -> string SVG listo para insertar
//   iconEl('bell')                      -> elemento <span class="icon-x"> con el SVG dentro, listo para .appendChild
//
// Debe cargarse ANTES de dashboard.js.
// =====================================================================
const ICONS = {
  refresh:        '<svg viewBox="0 0 24 24" class="icon"><path d="M3 12a9 9 0 0 1 15.5-6.3L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-15.5 6.3L3 16"/><path d="M3 21v-5h5"/></svg>',
  clipboardList:  '<svg viewBox="0 0 24 24" class="icon"><rect x="6" y="4" width="12" height="17" rx="2"/><path d="M9 4V3a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1"/><path d="M9 10h6M9 13h6M9 16h4"/></svg>',
  zap:            '<svg viewBox="0 0 24 24" class="icon"><path d="M13 2 3 14h7l-1 8 10-12h-7z"/></svg>',
  barChart:       '<svg viewBox="0 0 24 24" class="icon"><path d="M4 20V10M12 20V4M20 20v-7"/></svg>',
  bell:           '<svg viewBox="0 0 24 24" class="icon"><path d="M6 9a6 6 0 0 1 12 0c0 5 2 6 2 6H4s2-1 2-6"/><path d="M10 20a2 2 0 0 0 4 0"/></svg>',
  users:          '<svg viewBox="0 0 24 24" class="icon"><circle cx="9" cy="8" r="3.2"/><path d="M2.5 20c0-3.5 3-6 6.5-6s6.5 2.5 6.5 6"/><path d="M16 8.2a3 3 0 0 1 0 5.9"/><path d="M17.5 14c2.8.4 4.5 2.6 4.5 6"/></svg>',
  unplug:         '<svg viewBox="0 0 24 24" class="icon"><path d="M9 3v5M15 3v5M6 8h12v3a6 6 0 0 1-12 0V8Z"/><path d="M12 16v5"/></svg>',
  moon:           '<svg viewBox="0 0 24 24" class="icon"><path d="M20 14.5A8.5 8.5 0 1 1 9.5 4a7 7 0 0 0 10.5 10.5Z"/></svg>',
  batteryWarning: '<svg viewBox="0 0 24 24" class="icon"><rect x="2" y="7" width="17" height="10" rx="2"/><path d="M22 10v4"/><path d="M9 9.5 7 13h3l-2 3.5"/></svg>',
  tag:            '<svg viewBox="0 0 24 24" class="icon"><path d="M20.6 12.6 12.6 20.6a2 2 0 0 1-2.8 0L3 13.8V5a2 2 0 0 1 2-2h8.8a2 2 0 0 1 1.4.6l7.4 7.4a2 2 0 0 1 0 2.6Z"/><circle cx="8" cy="8" r="1.3" fill="currentColor" stroke="none"/></svg>',
  clock:          '<svg viewBox="0 0 24 24" class="icon"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.5 2"/></svg>',
  trendingDown:   '<svg viewBox="0 0 24 24" class="icon"><path d="M3 6l7 7 4-4 7 7"/><path d="M15 16h6v-6"/></svg>',
  alarmClock:     '<svg viewBox="0 0 24 24" class="icon"><circle cx="12" cy="13" r="8"/><path d="M12 9v4l3 2"/><path d="M5 3 2 6M19 3l3 3"/></svg>',
  alertTriangle:  '<svg viewBox="0 0 24 24" class="icon"><path d="M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/><path d="M12 9v4"/><path d="M12 17h.01"/></svg>',
  download:       '<svg viewBox="0 0 24 24" class="icon"><path d="M12 3v12"/><path d="M7 10l5 5 5-5"/><path d="M4 19h16"/></svg>',
  package:        '<svg viewBox="0 0 24 24" class="icon"><path d="M21 8 12 3 3 8v8l9 5 9-5V8Z"/><path d="M3 8l9 5 9-5"/><path d="M12 13v8"/></svg>',
  checkCircle:    '<svg viewBox="0 0 24 24" class="icon"><circle cx="12" cy="12" r="9"/><path d="M8 12.5 11 15.5 16 9"/></svg>',
  hourglass:      '<svg viewBox="0 0 24 24" class="icon"><path d="M6 2h12M6 22h12"/><path d="M6 2c0 5 4 6.5 6 8 2-1.5 6-3 6-8"/><path d="M6 22c0-5 4-6.5 6-8 2 1.5 6 3 6 8"/></svg>',
  warehouse:      '<svg viewBox="0 0 24 24" class="icon"><path d="M3 21V10l9-6 9 6v11"/><path d="M3 21h18"/><path d="M8 21v-6h8v6"/></svg>',
  x:              '<svg viewBox="0 0 24 24" class="icon"><path d="M5 5l14 14M19 5 5 19"/></svg>'
};

// Escapa texto que va a insertarse como HTML (nombres de cliente, placas,
// distritos, etc. que vienen del Excel/CSV cargado). Previene que un valor
// con caracteres < > & " ' rompa el layout o se interprete como HTML.
function escapeHtml(value){
  if(value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
