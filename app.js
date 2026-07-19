/**
 * Conexión al backend (Apps Script). Reemplaza API_BASE con la URL de tu
 * Web App desplegada (termina en /exec) — está en Configuración →
 * URL_PORTAL_RESULTADOS dentro de tu Google Sheet.
 */
const API_BASE = 'https://script.google.com/macros/s/AKfycbzw0lAh-i0MvbSMJLvsBsG2pcjAx5q-PyGddtZAWlvBACwlSHnADOaRw7ER6FwJ0BD6/exec';

/**
 * Links de los formularios de REGISTRO (no de acceso a portal existente) —
 * para quien llega por primera vez y todavía no tiene un código. Cópialos
 * desde el menú "🔗 Ver links de formularios" en tu Google Sheet.
 */
const FORM_COMPRADOR_URL = 'https://forms.gle/W2qLDtgjEopHpF6Q6';
const FORM_AGENTE_URL = 'https://forms.gle/d6C866rynvJRXP2HA';
const FORM_EMBAJADOR_URL = 'https://forms.gle/hmw8EruUmRW7CynC8';

/** Número de WhatsApp de soporte (formato: 18095551234, con código de país) */
const WHATSAPP_SOPORTE_NUMERO = '18098012075';


function obtenerParametro(nombre) {
  return new URLSearchParams(window.location.search).get(nombre) || '';
}

async function llamarApi(params) {
  const url = API_BASE + '?api=1&' + new URLSearchParams(params).toString();
  const resp = await fetch(url);
  if (!resp.ok) throw new Error('Error de red: ' + resp.status);
  return resp.json();
}

async function ejecutarAccion(params) {
  const url = API_BASE + '?api=1&' + new URLSearchParams(params).toString();
  const resp = await fetch(url);
  return resp.json();
}

function guardarCodigoRecordado(tipo, codigo) {
  try { localStorage.setItem('propmatch_' + tipo, codigo); } catch (e) { /* si el navegador bloquea localStorage, simplemente no se recuerda */ }
}

function obtenerCodigoRecordado(tipo) {
  try { return localStorage.getItem('propmatch_' + tipo) || ''; } catch (e) { return ''; }
}

function mostrarError(contenedorId, mensaje) {
  document.getElementById(contenedorId).innerHTML =
    '<div class="card"><h3>No pudimos cargar esto</h3><p>' + mensaje + '</p></div>';
}

function construirRutaHtml(etapaActual, etapas) {
  const indiceActual = etapas.indexOf(etapaActual);
  return etapas.map((etapa, i) => {
    const clase = i < indiceActual ? 'hecho' : i === indiceActual ? 'actual' : '';
    const icono = i < indiceActual ? '✓' : i === indiceActual ? '●' : '○';
    return '<div class="ruta-paso ' + clase + '"><div class="ruta-nodo">' + icono + '</div>' +
      '<div class="ruta-texto">' + etapa + '</div></div>';
  }).join('');
}

function formatoRD(numero) {
  return 'RD$' + Math.round(numero).toLocaleString('es-DO');
}

/**
 * "Consultar mi solicitud" unificado: prueba el código contra el Portal del
 * Comprador primero, y si no existe ahí, contra el Panel del Agente — así
 * la persona no necesita saber de antemano "en qué categoría está", solo
 * pega su código una vez.
 */
async function consultarSolicitud(inputId, mensajeId) {
  const codigo = document.getElementById(inputId).value.trim();
  const msj = document.getElementById(mensajeId);
  if (!codigo) return;
  msj.textContent = 'Buscando tu solicitud…';

  try {
    const comoComprador = await llamarApi({ portal: codigo });
    if (!comoComprador.error) {
      guardarCodigoRecordado('comprador', codigo);
      window.location.href = 'portal.html?codigo=' + encodeURIComponent(codigo);
      return;
    }
    const comoAgente = await llamarApi({ panel: codigo });
    if (!comoAgente.error) {
      guardarCodigoRecordado('agente', codigo);
      window.location.href = 'panel.html?codigo=' + encodeURIComponent(codigo);
      return;
    }
    const comoEmbajador = await llamarApi({ embajador: codigo });
    if (!comoEmbajador.error) {
      window.location.href = 'panel-embajador.html?codigo=' + encodeURIComponent(codigo);
      return;
    }
    msj.textContent = '⚠️ No encontramos ninguna solicitud con ese código. Verifica el enlace que recibiste por correo.';
  } catch (e) {
    msj.textContent = '⚠️ Ocurrió un problema de conexión. Intenta de nuevo en unos minutos.';
  }
}

function construirBadgeClasificacion(clasificacion) {
  const iconos = { Diamante: '💎', Oro: '🥇', Plata: '🥈', Bronce: '🥉' };
  const clase = String(clasificacion || '').toLowerCase();
  const icono = iconos[clasificacion] || '⭐';
  return '<div class="badge-clasificacion ' + clase + '"><span class="icono">' + icono + '</span> ' + clasificacion + '</div>';
}

/**
 * Gráfico de gauge circular (dona) — para cualquier puntaje 0-100. Hecho en
 * SVG puro, sin librerías externas (gratis, cero dependencias, carga
 * instantánea).
 */
function construirGraficoGauge(porcentaje, color, tamano) {
  tamano = tamano || 120;
  color = color || '#0f7a6b';
  const radio = tamano / 2 - 10;
  const circunferencia = 2 * Math.PI * radio;
  const pct = Math.max(0, Math.min(100, porcentaje));
  const offset = circunferencia * (1 - pct / 100);
  return '<svg width="' + tamano + '" height="' + tamano + '" viewBox="0 0 ' + tamano + ' ' + tamano + '">' +
    '<circle cx="' + (tamano / 2) + '" cy="' + (tamano / 2) + '" r="' + radio + '" fill="none" stroke="#e9e0cc" stroke-width="10"/>' +
    '<circle cx="' + (tamano / 2) + '" cy="' + (tamano / 2) + '" r="' + radio + '" fill="none" stroke="' + color + '" stroke-width="10" ' +
    'stroke-dasharray="' + circunferencia + '" stroke-dashoffset="' + offset + '" stroke-linecap="round" ' +
    'transform="rotate(-90 ' + (tamano / 2) + ' ' + (tamano / 2) + ')"/>' +
    '<text x="50%" y="50%" text-anchor="middle" dy="0.35em" font-family="IBM Plex Mono, monospace" font-size="' + Math.round(tamano * 0.2) + '" font-weight="700" fill="#10231c">' + Math.round(pct) + '</text>' +
    '</svg>';
}

/** Gráfico de barras horizontales — para funnels (embajador) o comparativas simples */
function construirGraficoBarras(pasos) {
  const max = Math.max.apply(null, pasos.map(p => p.valor).concat([1]));
  return pasos.map(p => {
    const ancho = Math.max(6, Math.round((p.valor / max) * 100));
    return '<div style="margin-bottom:12px">' +
      '<div style="font-size:12px;color:var(--text-muted);margin-bottom:4px">' + p.etiqueta + ': <b style="color:var(--ink)">' + p.valor + '</b></div>' +
      '<div style="background:var(--sand-deep);border-radius:6px;overflow:hidden;height:16px">' +
      '<div style="background:' + (p.color || 'var(--teal)') + ';width:' + ancho + '%;height:100%;transition:width 0.4s ease"></div></div></div>';
  }).join('');
}
