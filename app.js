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
const FORM_COMPRADOR_URL = 'https://forms.gle/Ekv3n3hTe5rp3oF36';
const FORM_AGENTE_URL = 'https://forms.gle/JaSotmYaDvvQ2byD8';
const FORM_EMBAJADOR_URL = 'https://forms.gle/6epFrzpFsUYPrAqf7';
const FORM_ALQUILER_URL = 'https://forms.gle/XqmJ5fi1ePbReHdT6';
const FORM_PROPIETARIO_URL = 'https://forms.gle/A4LBpHchrC5qWTt66';

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

/**
 * Convierte un archivo (input type="file") a base64 y lo envía junto con
 * el resto de la acción, vía POST — gratis, sin ningún servicio de pago,
 * sin exigir que la persona inicie sesión en nada.
 */
/**
 * Comprime una foto antes de subirla — una foto directa de cámara de
 * celular puede pesar 3-8 MB, lo cual puede hacer fallar la subida (timeout
 * o límite de tamaño de la solicitud). Se reduce a un tamaño razonable
 * (1600px de lado más largo, calidad 80%) sin que el agente note ninguna
 * diferencia visual relevante para un contrato/evidencia. Los PDF no se
 * tocan, ya que no se pueden recomprimir de esta forma.
 */
function comprimirImagenSiAplica_(archivo) {
  if (!archivo.type.startsWith('image/')) return Promise.resolve(archivo);
  return new Promise((resolve) => {
    const img = new Image();
    const url = URL.createObjectURL(archivo);
    img.onload = () => {
      const MAX_LADO = 1600;
      let { width, height } = img;
      if (width > MAX_LADO || height > MAX_LADO) {
        const ratio = Math.min(MAX_LADO / width, MAX_LADO / height);
        width = Math.round(width * ratio);
        height = Math.round(height * ratio);
      }
      const canvas = document.createElement('canvas');
      canvas.width = width;
      canvas.height = height;
      canvas.getContext('2d').drawImage(img, 0, 0, width, height);
      canvas.toBlob((blob) => {
        URL.revokeObjectURL(url);
        resolve(blob ? new File([blob], archivo.name, { type: 'image/jpeg' }) : archivo);
      }, 'image/jpeg', 0.8);
    };
    img.onerror = () => { URL.revokeObjectURL(url); resolve(archivo); }; // si falla, se sube el original sin comprimir
    img.src = url;
  });
}

async function ejecutarAccionConArchivo(params, archivoInputElement) {
  let archivo = archivoInputElement && archivoInputElement.files && archivoInputElement.files[0];
  const payload = Object.assign({ api: '1' }, params);

  if (archivo) {
    archivo = await comprimirImagenSiAplica_(archivo);
    const base64 = await new Promise((resolve, reject) => {
      const lector = new FileReader();
      lector.onload = () => resolve(lector.result.split(',')[1]); // quita el prefijo "data:image/...;base64,"
      lector.onerror = reject;
      lector.readAsDataURL(archivo);
    });
    payload.archivoBase64 = base64;
    payload.archivoMime = archivo.type;
    payload.archivoNombre = archivo.name;
  }

  const resp = await fetch(API_BASE, { method: 'POST', body: JSON.stringify(payload) });
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
    const comoAlquiler = await llamarApi({ portalAlquiler: codigo });
    if (!comoAlquiler.error) {
      window.location.href = 'portal-alquiler.html?codigo=' + encodeURIComponent(codigo);
      return;
    }
    const comoPropietario = await llamarApi({ portalPropietario: codigo });
    if (!comoPropietario.error) {
      window.location.href = 'portal-propietario.html?codigo=' + encodeURIComponent(codigo);
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

/**
 * VULNERABILIDAD DE SEGURIDAD ENCONTRADA Y CORREGIDA: los nombres y textos
 * que un comprador/agente/inquilino escribe en el formulario público (sin
 * ninguna cuenta ni verificación) se mostraban directamente dentro del
 * innerHTML de los Paneles — si alguien registraba su nombre con código
 * malicioso (ej. una etiqueta <img> con un manejador de error), se
 * ejecutaría en el navegador de quien vea esa tarjeta (agente, admin).
 * Esta función escapa cualquier texto antes de insertarlo, convirtiendo
 * caracteres especiales de HTML en su forma segura — usar SIEMPRE que se
 * muestre un nombre, nota, o cualquier texto que haya escrito un usuario.
 */
function escaparHtml(texto) {
  if (texto === null || texto === undefined) return '';
  return String(texto)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Gráfico simple (SVG, sin librerías externas) que muestra el presupuesto
 * de un comprador frente al rango real de precios del mercado en su zona
 * y tipo de inmueble — de un vistazo, sin tener que leer un párrafo de
 * texto para entender si el presupuesto es realista.
 */
function construirGraficoPresupuesto(g) {
  if (!g || (!g.mercadoMin && !g.mercadoMax)) {
    return '<p style="font-size:13px;color:var(--text-muted)">Sin referencia de mercado cargada todavía para esta zona/tipo de inmueble.</p>';
  }
  const anchoTotal = 340, alto = 100;
  const formato = v => 'RD$' + Math.round(v).toLocaleString('es-DO');

  // El rango visible incluye tanto el mercado como el presupuesto, con
  // un 15% de margen a cada lado para que nada quede pegado al borde.
  const valores = [g.mercadoMin, g.mercadoMax, g.presupuestoMax].filter(v => v > 0);
  const min = Math.min(...valores) * 0.85;
  const max = Math.max(...valores) * 1.15;
  const escala = v => 12 + ((v - min) / (max - min)) * (anchoTotal - 24);

  const barraX1 = escala(g.mercadoMin), barraX2 = escala(g.mercadoMax);
  const marcadorX = Math.max(24, Math.min(anchoTotal - 24, escala(g.presupuestoMax)));
  const dentroDelRango = g.presupuestoMax >= g.mercadoMin && g.presupuestoMax <= g.mercadoMax * 1.05;
  const colorMarcador = dentroDelRango ? '#0f7a6b' : (g.presupuestoMax < g.mercadoMin ? '#e4572e' : '#2e6b8a');

  return '<svg viewBox="0 0 ' + anchoTotal + ' ' + alto + '" style="width:100%;max-width:400px;height:auto" role="img" aria-label="Presupuesto comparado con el mercado">' +
    // Etiqueta superior de la barra de mercado
    '<text x="' + ((barraX1 + barraX2) / 2) + '" y="14" font-size="10" fill="#8a8272" text-anchor="middle">Rango de precios en tu zona</text>' +
    // Riel de fondo + barra de mercado
    '<line x1="12" y1="34" x2="' + (anchoTotal - 12) + '" y2="34" stroke="#e4dcc8" stroke-width="8" stroke-linecap="round"/>' +
    '<line x1="' + barraX1 + '" y1="34" x2="' + barraX2 + '" y2="34" stroke="#c9bfa0" stroke-width="8" stroke-linecap="round"/>' +
    // Montos del rango de mercado, debajo de la barra
    '<text x="' + barraX1 + '" y="52" font-size="11" font-weight="600" fill="#6b6552" text-anchor="middle">' + formato(g.mercadoMin) + '</text>' +
    '<text x="' + barraX2 + '" y="52" font-size="11" font-weight="600" fill="#6b6552" text-anchor="middle">' + formato(g.mercadoMax) + '</text>' +
    // Marcador del presupuesto
    '<circle cx="' + marcadorX + '" cy="34" r="8" fill="' + colorMarcador + '" stroke="white" stroke-width="2.5"/>' +
    '<line x1="' + marcadorX + '" y1="42" x2="' + marcadorX + '" y2="70" stroke="' + colorMarcador + '" stroke-width="1.5" stroke-dasharray="2,2"/>' +
    '<text x="' + marcadorX + '" y="84" font-size="12" font-weight="700" fill="' + colorMarcador + '" text-anchor="middle">Tu presupuesto</text>' +
    '<text x="' + marcadorX + '" y="98" font-size="13" font-weight="700" fill="' + colorMarcador + '" text-anchor="middle">' + formato(g.presupuestoMax) + '</text>' +
    '</svg>';
}

/**
 * Esqueleto de carga (shimmer) reutilizable en todos los paneles/portales —
 * reemplaza el spinner genérico por la forma aproximada del contenido que
 * está por llegar (estadísticas + tarjetas de lista), se percibe más
 * rápido e intencional que un círculo girando.
 */
function construirEsqueletoCarga(numTarjetas) {
  numTarjetas = numTarjetas || 3;
  const statsEsqueleto = ['', '', '', ''].map(() => '<div class="skeleton-shimmer skeleton-stat"></div>').join('');
  const tarjetasEsqueleto = Array.from({ length: numTarjetas }).map(() =>
    '<div class="skeleton-shimmer skeleton-card"></div>'
  ).join('');
  return '<div class="stat-grid" style="margin-bottom:16px">' + statsEsqueleto + '</div>' +
    '<div class="skeleton-shimmer skeleton-line" style="width:60%;height:22px;margin:0 auto 16px"></div>' +
    tarjetasEsqueleto;
}

/**
 * ============================================================================
 *  MODO MANTENIMIENTO — un solo interruptor centralizado para TODOS los
 *  Portales/Paneles (no solo la página principal). Cámbialo a "true"
 *  antes de aplicar cambios grandes, sube este archivo, y cualquier
 *  Portal/Panel que lo use mostrará el mensaje de mantenimiento en vez
 *  de su contenido normal — no depende del backend, funciona incluso si
 *  Apps Script está roto en ese momento.
 * ============================================================================
 */
const MODO_MANTENIMIENTO = false;

/**
 * Llamar al inicio de cada Portal/Panel — si el modo mantenimiento está
 * activo, reemplaza el contenido por el mensaje y devuelve true (para
 * que la página sepa detenerse ahí, sin intentar cargar datos del
 * backend).
 */
function mostrarMantenimientoSiAplica(idContenedor) {
  if (!MODO_MANTENIMIENTO) return false;
  document.getElementById(idContenedor).innerHTML =
    '<div style="text-align:center;padding-top:60px">' +
    '<span style="font-size:56px">🛠️</span>' +
    '<h1 style="margin-top:16px">Estamos mejorando PropMatchRD</h1>' +
    '<p style="font-size:16px;color:var(--text-muted);max-width:420px;margin:0 auto">Volvemos en breve. Gracias por tu paciencia.</p>' +
    '</div>';
  return true;
}

/**
 * Sección plegable reutilizable — colapsada por defecto, con un "Ver
 * más ▾" que la expande. Se usa para contenido secundario (gráficos,
 * historiales detallados) que no hace falta ver de entrada, para que
 * las tarjetas no se sientan tan densas cuando se apilan varios
 * bloques uno tras otro.
 */
function construirSeccionPlegable(idUnico, textoBoton, contenidoHtml) {
  return '<p style="font-size:13px;color:var(--teal-deep);cursor:pointer;margin:6px 0;user-select:none" onclick="toggleSeccionPlegable(\'' + idUnico + '\')">' +
    '<span id="' + idUnico + '_texto">' + escaparHtml(textoBoton) + '</span> ' +
    '<span id="' + idUnico + '_flecha" style="display:inline-block;transition:transform 0.15s">▾</span>' +
    '</p>' +
    '<div id="' + idUnico + '" style="display:none">' + contenidoHtml + '</div>';
}

function toggleSeccionPlegable(idUnico) {
  const contenido = document.getElementById(idUnico);
  const flecha = document.getElementById(idUnico + '_flecha');
  const abierto = contenido.style.display !== 'none';
  contenido.style.display = abierto ? 'none' : 'block';
  if (flecha) flecha.style.transform = abierto ? 'rotate(0deg)' : 'rotate(180deg)';
}

/**
 * Lenguaje de color unificado para "Compatibilidad" — antes el fondo de
 * la tarjeta era siempre el mismo color neutro sin importar si era Alta,
 * Media o Baja, a pesar de que el emoji sí cambiaba (🟢🟡🔴). Ahora el
 * fondo también refleja el nivel, para que se pueda "leer" de un
 * vistazo sin tener que leer el texto — mismo sistema en Panel y Portal.
 */
function colorFondoPorNivelCompatibilidad(nivel) {
  if (nivel === 'Alta') return '#e6f4f1';   // tinte del teal — positivo
  if (nivel === 'Media') return '#fdf3e0';  // tinte de ámbar — atención
  return '#fbe9e5';                          // tinte de coral — alerta
}
