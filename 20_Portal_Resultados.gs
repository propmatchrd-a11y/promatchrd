/**
 * ============================================================================
 *  PORTAL DE RESULTADOS (Web App gratis con Apps Script)
 * ============================================================================
 *  Google Forms NO puede mostrar una pantalla de confirmación dinámica — el
 *  mensaje de confirmación es texto fijo, igual para todos. Por eso, la
 *  verdadera "pantalla interactiva con tu análisis" solo es posible con una
 *  página web aparte. Esta es esa página: se despliega como Web App de
 *  Apps Script (gratis, sin servidor propio) y se linkea desde el correo de
 *  confirmación — el comprador hace clic y ve su análisis personalizado en
 *  una página, no solo en texto de correo.
 *
 *  IMPORTANTE — paso manual único de activación (ver manual de instalación):
 *  1) En el editor de Apps Script: Implementar → Nueva implementación.
 *  2) Tipo: Aplicación web.
 *  3) Ejecutar como: Yo (el propietario).
 *  4) Quién tiene acceso: Cualquier usuario.
 *  5) Implementar → copiar la URL que te da Google.
 *  6) Pegar esa URL en Configuración → URL_PORTAL_RESULTADOS.
 *  Sin este paso, el sistema sigue funcionando normal — solo no se incluye
 *  el botón "Ver mi análisis" en los correos.
 * ============================================================================
 */

/**
 * Reutiliza el análisis personalizado ya guardado en el registro (evita
 * llamar a Gemini de nuevo en cada carga del Portal — eso era lo que hacía
 * la carga lenta). Auto-repara compradores antiguos que se registraron
 * antes de que este campo existiera: lo calcula UNA vez y lo guarda, para
 * que la próxima carga ya sea instantánea también.
 */
function obtenerAnalisisPersonalizadoConCache_(obj, faltantes, alternativas, fila) {
  if (obj.Analisis_Personalizado_Texto) return obj.Analisis_Personalizado_Texto;
  const analisis = generarAnalisisPersonalizado_(obj, faltantes, alternativas);
  if (fila > -1) {
    sh_(SHEETS.COMPRADORES).getRange(fila, col_(SHEETS.COMPRADORES, 'Analisis_Personalizado_Texto')).setValue(analisis);
  }
  return analisis;
}

/**
 * Recibe solicitudes POST — usado ÚNICAMENTE para adjuntar evidencia
 * (foto/PDF del contrato) desde el Panel de Agente. Se eligió este camino
 * en vez de la subida nativa de Google Forms porque esa opción exige que
 * quien llena el formulario inicie sesión con una cuenta de Google — con
 * este método, cualquiera puede adjuntar una foto sin iniciar sesión en
 * nada, gratis, sin ningún servicio de pago de por medio.
 */
function doPost(e) {
  try {
    const datos = JSON.parse(e.postData.contents);
    let evidenciaUrl = '';

    if (datos.archivoBase64) {
      evidenciaUrl = subirEvidenciaComoArchivo_(datos.archivoBase64, datos.archivoMime, datos.archivoNombre);
    }

    // Se construye un objeto "e" compatible con lo que ya espera
    // ejecutarAccionMantenimiento_ — mismo despachador que usa doGet, sin
    // duplicar ninguna lógica de negocio.
    const eCompatible = { parameter: Object.assign({}, datos, { evidenciaUrl: evidenciaUrl || datos.evidenciaUrl || '' }) };
    const resultado = ejecutarAccionMantenimiento_(eCompatible, datos.accion);
    return ContentService.createTextOutput(JSON.stringify(resultado)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    log_('doPost', 'ERROR', err.message + ' | ' + err.stack);
    return ContentService.createTextOutput(JSON.stringify({ error: true, mensaje: 'Ocurrió un error procesando el archivo.' })).setMimeType(ContentService.MimeType.JSON);
  }
}

/** Sube un archivo (foto/PDF) recibido en base64 a Drive, gratis, y devuelve su link público de solo lectura */
function subirEvidenciaComoArchivo_(base64Data, mimeType, nombreArchivo) {
  try {
    const props = PropertiesService.getScriptProperties();
    let folderId = props.getProperty('CARPETA_EVIDENCIAS_ID');
    let carpeta;
    if (folderId) {
      try { carpeta = DriveApp.getFolderById(folderId); } catch (e) { /* recrear abajo */ }
    }
    if (!carpeta) {
      const nombreCarpeta = APP_NAME + ' — Evidencias de Cierre';
      const existentes = DriveApp.getFoldersByName(nombreCarpeta);
      carpeta = existentes.hasNext() ? existentes.next() : DriveApp.createFolder(nombreCarpeta);
      props.setProperty('CARPETA_EVIDENCIAS_ID', carpeta.getId());
    }

    const bytes = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', nombreArchivo || ('evidencia_' + new Date().getTime()));
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return archivo.getUrl();
  } catch (err) {
    log_('subirEvidenciaComoArchivo_', 'ERROR', err.message + ' | ' + err.stack);
    return '';
  }
}

function doGet(e) {
  try {
    // Confirmación de monto de cierre (clic directo desde el correo del
    // comprador) — se maneja ANTES del modo API porque este link se abre
    // directo en el navegador (no por fetch/JS), así que debe mostrar una
    // página clara y humana, no una respuesta técnica en JSON.
    if (e.parameter.accion === 'confirmarMontoCierre') {
      return construirPaginaConfirmacionMonto_(e);
    }

    // MODO API (JSON): usado por el frontend de GitHub Pages. En vez de
    // devolver una página HTML completa, devuelve solo los datos — el
    // diseño visual vive en GitHub Pages, este backend solo entrega la
    // información. Se activa con ?api=1 junto al parámetro que corresponda.
    if (e.parameter.api === '1') {
      return manejarSolicitudApi_(e);
    }

    // Landing page de Red de Embajadores: ?ref=CODIGO
    const ref = (e.parameter.ref || '').toUpperCase().trim();
    if (ref) {
      return HtmlService.createHtmlOutput(construirLandingEmbajador_(ref))
        .setTitle('Te invitaron a ' + APP_NAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // Dashboard privado del Agente: ?panel=CODIGO_PRIVADO (NO el ID público —
    // ver generarCodigoPrivadoLargo_ en 01_Utils.gs). Con &contacto=CVRD-XXX
    // marca el primer contacto de ese comprador — así el SLA se alimenta
    // directo del agente con un clic, sin que el admin tenga que hacerlo.
    const codigoPanel = (e.parameter.panel || '').trim();
    if (codigoPanel) {
      const idComprobarContacto = (e.parameter.contacto || '').toUpperCase().trim();
      return HtmlService.createHtmlOutput(construirPanelAgente_(codigoPanel, idComprobarContacto))
        .setTitle('Tu panel — ' + APP_NAME)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    }

    // Portal de Resultados del comprador: ?portal=CODIGO_PRIVADO (NO el ID
    // público CVRD-XXXXXX — usar el ID secuencial habría permitido que
    // cualquiera probara números consecutivos y viera información financiera
    // privada de otros compradores. Mismo criterio de seguridad que el
    // Panel del Agente — ver generarCodigoPrivadoLargo_ en 01_Utils.gs).
    const codigoPortal = (e.parameter.portal || '').trim();
    if (!codigoPortal) {
      return HtmlService.createHtmlOutput(paginaBase_(
        '<h2>PropMatchRD</h2><p>Falta el código de acceso en el enlace.</p>'
      ));
    }

    const fila = buscarFilaPorCodigoPortal_(codigoPortal);
    if (fila === -1) {
      return HtmlService.createHtmlOutput(paginaBase_(
        '<h2>No encontramos ese código</h2><p>Verifica el enlace que recibiste por correo.</p>'
      ));
    }

    const obj = filaAObjeto_(SHEETS.COMPRADORES, sh_(SHEETS.COMPRADORES).getRange(fila, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
    const { faltantes } = calcularIndicePreparacion_(obj);
    const alternativas = obtenerProvinciasAlternativasSimple_(obj.Presupuesto_Max, obj.Provincia);
    const analisis = obtenerAnalisisPersonalizadoConCache_(obj, faltantes, alternativas, fila);

    const colorClasif = { Diamante: '#0b5394', Oro: '#7f6000', Plata: '#434343', Bronce: '#783f04' }[obj.Clasificacion] || '#1a73e8';
    const bgClasif = { Diamante: '#d0e8ff', Oro: '#fff2cc', Plata: '#f3f3f3', Bronce: '#fce5cd' }[obj.Clasificacion] || '#e8f0fe';

    const listaPasos = faltantes.length
      ? '<ul style="margin:8px 0;padding-left:20px">' + faltantes.map(f => '<li style="margin-bottom:6px">' + f + '</li>').join('') + '</ul>'
      : '<p>¡Tu perfil está completo! Un agente verificado se pondrá en contacto pronto.</p>';

    const contenido =
      '<div style="text-align:center;margin-bottom:20px">' +
      '<span style="background:' + bgClasif + ';color:' + colorClasif + ';padding:6px 16px;border-radius:20px;font-weight:bold;font-size:15px">' +
      obj.Clasificacion + '</span></div>' +
      '<h2 style="text-align:center;margin-top:0">Hola ' + obj.Nombre + ', este es tu análisis</h2>' +
      '<div style="background:#f0f7ff;padding:16px;border-radius:8px;margin:16px 0">' +
      '<p style="margin:0;font-size:16px;line-height:1.5">' + analisis + '</p></div>' +

      '<div style="display:flex;gap:12px;margin:16px 0;flex-wrap:wrap">' +
      '<div style="flex:1;min-width:140px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px;text-align:center">' +
      '<div style="font-size:26px;font-weight:bold;color:#1a73e8">' + obj.Indice_Preparacion_Pct + '%</div>' +
      '<div style="font-size:13px;color:#666">Índice de Preparación</div></div>' +
      '<div style="flex:1;min-width:140px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px;text-align:center">' +
      '<div style="font-size:26px;font-weight:bold;color:#1a73e8">' + obj.Probabilidad_Financiamiento_Pct + '%</div>' +
      '<div style="font-size:13px;color:#666">Probabilidad de Financiamiento</div></div>' +
      '</div>' +

      '<h3>Tu Ruta de Compra</h3>' + construirRutaCompraVisual_(obj) +

      '<h3>💰 Gastos a considerar (más allá del precio)</h3>' +
      '<div style="background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px;margin:12px 0">' +
      calcularGastosAdicionalesCompra_(obj).html + '</div>' +

      '<h3>¿Qué sigue?</h3>' + listaPasos +

      (obj.Alerta_Presupuesto ? '<div style="background:#fff3cd;padding:12px;border-radius:8px;margin:12px 0"><b>Sobre tu presupuesto:</b> ' + obj.Alerta_Presupuesto + '</div>' : '') +
      (obj.Certificado_URL ? '<p style="text-align:center;margin-top:20px"><a href="' + obj.Certificado_URL + '" style="background:#1a73e8;color:white;padding:10px 20px;border-radius:6px;text-decoration:none">📄 Ver mi Certificado de Comprador Verificado</a></p>' : '') +
      '<p style="text-align:center;color:#999;font-size:12px;margin-top:24px">Código: ' + obj.ID_Comprador + '</p>';

    return HtmlService.createHtmlOutput(paginaBase_(contenido))
      .setTitle('Tu análisis — ' + APP_NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    log_('doGet', 'ERROR', err.message + ' | ' + err.stack);
    return HtmlService.createHtmlOutput(paginaBase_('<p>Ocurrió un error mostrando tu análisis. Contáctanos por WhatsApp.</p>'));
  }
}

/** Construye el camino visual de las 8 etapas de la Ruta de Compra, con la etapa actual resaltada */
function construirRutaCompraVisual_(comprador) {
  const etapaActual = comprador.Etapa_Ruta_Compra || 'Evaluación';
  const indiceActual = ETAPAS_RUTA_COMPRA.indexOf(etapaActual);

  const pasos = ETAPAS_RUTA_COMPRA.map((etapa, i) => {
    let icono, color;
    if (i < indiceActual) { icono = '✅'; color = '#274e13'; }
    else if (i === indiceActual) { icono = '🔵'; color = '#0b5394'; }
    else { icono = '⬜'; color = '#999'; }
    return '<div style="display:flex;align-items:center;padding:6px 0;color:' + color + (i === indiceActual ? ';font-weight:bold' : '') + '">' +
      '<span style="margin-right:8px">' + icono + '</span><span>' + etapa + '</span></div>';
  }).join('');

  return '<div style="background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:14px;margin:12px 0">' + pasos + '</div>';
}

/**
 * Dashboard del Agente (?panel=CODIGO_PRIVADO) — lo que faltaba: antes los
 * agentes solo recibían correos puntuales, sin ninguna vista propia de su
 * cartera completa. Muestra sus compradores activos (con su Etapa de Ruta
 * de Compra e Índice de Probabilidad de Cierre), KPIs de desempeño, su
 * nivel de gamificación, y si tiene alguna comisión pendiente de pago.
 */
function construirPanelAgente_(codigoPanel, idComprobarContacto) {
  const filaAgente = buscarFilaAgentePorCodigoPanel_(codigoPanel);
  if (filaAgente === -1) {
    return paginaBase_('<h2>Enlace no válido</h2><p>Verifica que copiaste el link completo (revisa que no se haya cortado al copiarlo), o contáctanos si crees que es un error.</p>');
  }

  const agente = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);

  let confirmacionContacto = '';
  if (idComprobarContacto) {
    confirmacionContacto = marcarPrimerContactoDesdeAgente_(idComprobarContacto, agente.ID_Agente);
  }

  const compradoresSheet = sh_(SHEETS.COMPRADORES);
  const lastRowComp = compradoresSheet.getLastRow();
  const misCompradores = lastRowComp > 1
    ? compradoresSheet.getRange(2, 1, lastRowComp - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
        .filter(c => c.Agente_Asignado_ID === agente.ID_Agente && !['Cerrado', 'Perdido', 'Duplicado'].includes(c.Estado))
    : [];

  const cierresSheet = sh_(SHEETS.CIERRES);
  const lastRowCierres = cierresSheet.getLastRow();
  const comisionesPendientes = lastRowCierres > 1
    ? cierresSheet.getRange(2, 1, lastRowCierres - 1, HEADERS[SHEETS.CIERRES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CIERRES, f))
        .filter(c => c.ID_Agente === agente.ID_Agente && c.Estado_Pago_Comision !== 'Pagado')
    : [];

  // Mapa de contactos ya marcados, para no repetir el botón innecesariamente
  const asignacionesSheet = sh_(SHEETS.ASIGNACIONES);
  const lastRowAsig = asignacionesSheet.getLastRow();
  const contactosPendientes = {};
  if (lastRowAsig > 1) {
    asignacionesSheet.getRange(2, 1, lastRowAsig - 1, HEADERS[SHEETS.ASIGNACIONES].length).getValues()
      .map(f => filaAObjeto_(SHEETS.ASIGNACIONES, f))
      .filter(a => a.ID_Agente === agente.ID_Agente)
      .forEach(a => { contactosPendientes[a.ID_Comprador] = !a.Fecha_Primer_Contacto; });
  }

  const filasCompradores = misCompradores.map(c => {
    const botonContacto = contactosPendientes[c.ID_Comprador]
      ? '<a href="?panel=' + encodeURIComponent(codigoPanel) + '&contacto=' + encodeURIComponent(c.ID_Comprador) +
        '" style="background:#1a73e8;color:white;padding:4px 10px;border-radius:4px;text-decoration:none;font-size:11px">Marcar contacto</a>'
      : '✅';
    return '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee">' + c.ID_Comprador + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee">' + c.Nombre + ' ' + c.Apellido + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee">' + (c.Etapa_Ruta_Compra || '—') + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee">' + (c.IPC_Etiqueta || '—') + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px solid #eee">' + botonContacto + '</td></tr>';
  }).join('');

  const contenido =
    confirmacionContacto +
    '<div style="text-align:center;margin-bottom:16px">' +
    '<span style="background:#e8f0fe;color:#0b5394;padding:6px 16px;border-radius:20px;font-weight:bold">' + (agente.Distintivo || '🥉 Bronze') + '</span>' +
    '</div>' +
    '<h2 style="text-align:center;margin-top:0">Hola ' + agente.Nombre + '</h2>' +

    '<div style="display:flex;gap:10px;margin:16px 0;flex-wrap:wrap">' +
    '<div style="flex:1;min-width:100px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold;color:#1a73e8">' + misCompradores.length + '</div>' +
    '<div style="font-size:11px;color:#666">Compradores activos</div></div>' +
    '<div style="flex:1;min-width:100px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold;color:#1a73e8">' + (agente.Cierres_Confirmados || 0) + '</div>' +
    '<div style="font-size:11px;color:#666">Cierres confirmados</div></div>' +
    '<div style="flex:1;min-width:100px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold;color:#1a73e8">' + (agente.Tasa_Conversion || 0) + '%</div>' +
    '<div style="font-size:11px;color:#666">Tasa de conversión</div></div>' +
    '<div style="flex:1;min-width:100px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;padding:12px;text-align:center">' +
    '<div style="font-size:22px;font-weight:bold;color:#1a73e8">' + (agente.Tiempo_Respuesta_Promedio_Horas || '—') + '</div>' +
    '<div style="font-size:11px;color:#666">Horas resp. promedio</div></div>' +
    '</div>' +

    (comisionesPendientes.length
      ? '<div style="background:#fff3cd;padding:10px;border-radius:8px;margin:12px 0">⚠️ Tienes <b>' + comisionesPendientes.length +
        '</b> comisión(es) pendiente(s) de pago. Revisa tu correo para coordinar.</div>'
      : '') +

    '<h3>Tus compradores activos</h3>' +
    (misCompradores.length
      ? '<table style="width:100%;border-collapse:collapse;font-size:13px"><tr style="background:#f8f9fa;font-weight:bold">' +
        '<td style="padding:6px 8px">ID</td><td style="padding:6px 8px">Nombre</td><td style="padding:6px 8px">Ruta</td><td style="padding:6px 8px">IPC</td><td style="padding:6px 8px">Contacto</td></tr>' +
        filasCompradores + '</table>'
      : '<p style="color:#666">No tienes compradores activos en este momento.</p>');

  return paginaBase_(contenido);
}

/** Busca la fila (1-based) de un agente por su código privado de panel, o -1 si no existe */
/**
 * Marca Fecha_Primer_Contacto en "Asignaciones" cuando el AGENTE lo confirma
 * desde su propio Panel — así el SLA se alimenta directo de quien hace el
 * trabajo, sin depender de que el admin lo actualice manualmente en la hoja.
 */
function marcarPrimerContactoDesdeAgente_(idComprador, idAgente) {
  const sheet = sh_(SHEETS.ASIGNACIONES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const headers = HEADERS[SHEETS.ASIGNACIONES];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const idxComprador = headers.indexOf('ID_Comprador');
  const idxAgente = headers.indexOf('ID_Agente');
  const idxFechaContacto = headers.indexOf('Fecha_Primer_Contacto');

  for (let i = 0; i < data.length; i++) {
    if (data[i][idxComprador] === idComprador && data[i][idxAgente] === idAgente) {
      if (!data[i][idxFechaContacto]) {
        sheet.getRange(i + 2, idxFechaContacto + 1).setValue(new Date());
        bitacora_('Agente', 'Primer contacto marcado desde el Panel', idComprador, idAgente);
        return '<div style="background:#d9ead3;padding:10px;border-radius:8px;margin-bottom:12px">✅ Primer contacto con ' + idComprador + ' registrado.</div>';
      }
      return '<div style="background:#e8f0fe;padding:10px;border-radius:8px;margin-bottom:12px">Ya tenías registrado el primer contacto con ' + idComprador + '.</div>';
    }
  }
  return '';
}

function buscarFilaAgentePorCodigoPanel_(codigo) {
  const sheet = sh_(SHEETS.AGENTES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const codigoBuscado = String(codigo).trim();
  const codigos = sheet.getRange(2, col_(SHEETS.AGENTES, 'Codigo_Panel_Privado'), lastRow - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) return i + 2;
  }
  return -1;
}

/** Construye el link privado al Dashboard del Agente, si el Portal ya está desplegado */
function construirLinkPanelAgente_(codigoPanel) {
  if (!codigoPanel) return '';
  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  if (urlGithub) return urlGithub.replace(/\/$/, '') + '/panel.html?codigo=' + encodeURIComponent(codigoPanel);
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  if (!urlBase) return '';
  const separador = urlBase.includes('?') ? '&' : '?';
  return urlBase + separador + 'panel=' + encodeURIComponent(codigoPanel);
}

/**
 * Lógica compartida de confirmación de monto — usada tanto por el clic
 * directo del correo (página HTML clara) como por el modo API (JSON), para
 * no duplicar la misma lógica en dos lugares.
 */
function procesarConfirmacionMontoCierre_(compradorObj, idCierre, valor) {
  const filaCierre = buscarFilaPorId_(SHEETS.CIERRES, idCierre);
  if (filaCierre === -1) return { error: true, mensaje: 'No encontramos ese cierre. Es posible que el link haya expirado o esté incompleto.' };

  const cierresSheet = sh_(SHEETS.CIERRES);
  const cierreActual = filaAObjeto_(SHEETS.CIERRES, cierresSheet.getRange(filaCierre, 1, 1, HEADERS[SHEETS.CIERRES].length).getValues()[0]);
  if (cierreActual.Monto_Confirmado_Comprador && cierreActual.Monto_Confirmado_Comprador !== 'Pendiente') {
    return { error: false, mensaje: 'Ya habías respondido esto anteriormente — ¡gracias!' };
  }

  if (valor === 'No') {
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES, 'Requiere_Revision_Manual')).setValue('Sí');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES, 'Motivo_Revision')).setValue('El comprador indicó que el monto reportado NO es correcto.');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES, 'Estado_Pago_Comision')).setValue('En revisión');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES, 'Monto_Confirmado_Comprador')).setValue('No');
    const emailAdmin = getConfig_('EMAIL_ADMIN', '');
    if (emailAdmin) {
      enviarCorreo_(emailAdmin, '🚨 Comprador indica que el monto del cierre NO es correcto — ' + idCierre,
        '<p><b>' + idCierre + '</b> — el comprador ' + compradorObj.ID_Comprador + ' indicó que el monto reportado por el agente no coincide con lo acordado. Revisa antes de aprobar el pago.</p>');
    }
    bitacora_('Comprador', 'Monto de cierre marcado como incorrecto', idCierre, compradorObj.ID_Comprador);
    return { error: false, mensaje: 'Gracias por avisarnos — nuestro equipo revisará este cierre antes de procesar cualquier pago.' };
  }
  cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES, 'Monto_Confirmado_Comprador')).setValue('Sí');
  bitacora_('Comprador', 'Monto de cierre confirmado', idCierre, compradorObj.ID_Comprador);
  return { error: false, mensaje: '¡Gracias por confirmar!' };
}

/**
 * Página HTML clara y sencilla para el clic directo desde el correo del
 * comprador — antes esto mostraba una respuesta técnica en JSON, confusa
 * para cualquier persona que no sea desarrolladora. Ahora es una página
 * simple, humana, con el resultado explicado en una frase.
 */
function construirPaginaConfirmacionMonto_(e) {
  const codigoPortal = (e.parameter.portal || '').trim();
  const codigoPortalAlquiler = (e.parameter.portalAlquiler || '').trim();
  const idCierre = (e.parameter.idCierre || '').trim().toUpperCase();
  const valor = (e.parameter.valor || '').trim();

  let resultado;
  if (codigoPortalAlquiler) {
    const filaSolicitud = buscarFilaAlquilerPorCodigoPortal_(codigoPortalAlquiler);
    if (filaSolicitud === -1) {
      return HtmlService.createHtmlOutput(paginaBase_('<h2>Enlace no válido</h2><p>Verifica que copiaste el link completo desde tu correo, o contáctanos si crees que es un error.</p>'));
    }
    const solicitudObj = filaAObjeto_(SHEETS.SOLICITUDES_ALQUILER, sh_(SHEETS.SOLICITUDES_ALQUILER).getRange(filaSolicitud, 1, 1, HEADERS[SHEETS.SOLICITUDES_ALQUILER].length).getValues()[0]);
    resultado = procesarConfirmacionMontoCierreAlquiler_(solicitudObj, idCierre, valor);
  } else {
    const filaComprador = buscarFilaPorCodigoPortal_(codigoPortal);
    if (filaComprador === -1) {
      return HtmlService.createHtmlOutput(paginaBase_('<h2>Enlace no válido</h2><p>Verifica que copiaste el link completo desde tu correo, o contáctanos si crees que es un error.</p>'));
    }
    const compradorObj = filaAObjeto_(SHEETS.COMPRADORES, sh_(SHEETS.COMPRADORES).getRange(filaComprador, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
    resultado = procesarConfirmacionMontoCierre_(compradorObj, idCierre, valor);
  }

  const icono = resultado.error ? '⚠️' : (valor === 'No' ? '📋' : '✅');
  const colorFondo = resultado.error ? '#fce5cd' : (valor === 'No' ? '#fff3cd' : '#d9ead3');

  return HtmlService.createHtmlOutput(paginaBase_(
    '<div style="text-align:center;padding-top:40px">' +
    '<div style="font-size:48px">' + icono + '</div>' +
    '<div style="background:' + colorFondo + ';padding:20px;border-radius:10px;margin-top:16px">' +
    '<p style="font-size:17px;margin:0">' + resultado.mensaje + '</p>' +
    '</div>' +
    '<p style="margin-top:24px;color:#666;font-size:13px">Ya puedes cerrar esta ventana.</p>' +
    '</div>'
  )).setTitle('Confirmación — ' + APP_NAME);
}

/** Misma lógica que procesarConfirmacionMontoCierre_, para el módulo de Alquiler */
function procesarConfirmacionMontoCierreAlquiler_(solicitudObj, idCierre, valor) {
  const filaCierre = buscarFilaPorId_(SHEETS.CIERRES_ALQUILER, idCierre);
  if (filaCierre === -1) return { error: true, mensaje: 'No encontramos ese cierre. Es posible que el link haya expirado o esté incompleto.' };

  const cierresSheet = sh_(SHEETS.CIERRES_ALQUILER);
  const cierreActual = filaAObjeto_(SHEETS.CIERRES_ALQUILER, cierresSheet.getRange(filaCierre, 1, 1, HEADERS[SHEETS.CIERRES_ALQUILER].length).getValues()[0]);
  if (cierreActual.Monto_Confirmado_Comprador && cierreActual.Monto_Confirmado_Comprador !== 'Pendiente') {
    return { error: false, mensaje: 'Ya habías respondido esto anteriormente — ¡gracias!' };
  }

  if (valor === 'No') {
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES_ALQUILER, 'Requiere_Revision_Manual')).setValue('Sí');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES_ALQUILER, 'Motivo_Revision')).setValue('El inquilino indicó que el monto reportado NO es correcto.');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES_ALQUILER, 'Estado_Pago_Comision')).setValue('En revisión');
    cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES_ALQUILER, 'Monto_Confirmado_Comprador')).setValue('No');
    const emailAdmin = getConfig_('EMAIL_ADMIN', '');
    if (emailAdmin) {
      enviarCorreo_(emailAdmin, '🚨 Inquilino indica que el monto del cierre de alquiler NO es correcto — ' + idCierre,
        '<p><b>' + idCierre + '</b> — el inquilino de la solicitud ' + solicitudObj.ID_Solicitud + ' indicó que el monto reportado por el agente no coincide con lo acordado. Revisa antes de aprobar el pago.</p>');
    }
    bitacora_('Alquiler', 'Monto de cierre marcado como incorrecto', idCierre, solicitudObj.ID_Solicitud);
    return { error: false, mensaje: 'Gracias por avisarnos — nuestro equipo revisará este cierre antes de procesar cualquier pago.' };
  }
  cierresSheet.getRange(filaCierre, col_(SHEETS.CIERRES_ALQUILER, 'Monto_Confirmado_Comprador')).setValue('Sí');
  bitacora_('Alquiler', 'Monto de cierre confirmado', idCierre, solicitudObj.ID_Solicitud);
  return { error: false, mensaje: '¡Gracias por confirmar!' };
}

function paginaBase_(contenidoHtml) {
  return '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>body{font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:20px;background:#fafafa;color:#1a1a1a}' +
    'a{color:#1a73e8}</style></head><body>' + contenidoHtml + '</body></html>';
}

/** Construye el link privado al Portal de Resultados del comprador, si el Portal ya está desplegado */
function construirLinkPortalResultados_(codigoPortal) {
  if (!codigoPortal) return '';
  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  if (urlGithub) return urlGithub.replace(/\/$/, '') + '/portal.html?codigo=' + encodeURIComponent(codigoPortal);
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  if (!urlBase) return '';
  const separador = urlBase.includes('?') ? '&' : '?';
  return urlBase + separador + 'portal=' + encodeURIComponent(codigoPortal);
}

/** Busca la fila (1-based) de un comprador por su código privado de portal, o -1 si no existe */
function buscarFilaPorCodigoPortal_(codigo) {
  const sheet = sh_(SHEETS.COMPRADORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const codigoBuscado = String(codigo).trim();
  const codigos = sheet.getRange(2, col_(SHEETS.COMPRADORES, 'Codigo_Portal_Privado'), lastRow - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) return i + 2;
  }
  return -1;
}

/**
 * Construye el link del Formulario de Comprador con el código de embajador
 * YA prellenado en el campo "Código de Embajador (no modificar)" — así el
 * comprador nunca tiene que escribir a mano quién lo refirió.
 */
function construirLinkCompradorConReferido_(codigo) {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('FORM_COMPRADOR_ID');
  if (!formId) return '';
  try {
    const form = FormApp.openById(formId);
    const itemCodigo = form.getItems(FormApp.ItemType.TEXT)
      .find(it => it.getTitle() === 'Código de Embajador (no modificar)');
    if (!itemCodigo) return form.getPublishedUrl();
    const respuesta = form.createResponse();
    respuesta.withItemResponse(itemCodigo.asTextItem().createResponse(codigo));
    return respuesta.toPrefilledUrl();
  } catch (e) {
    return '';
  }
}

/** Mismo mecanismo que construirLinkCompradorConReferido_, para el formulario de Alquiler */
function construirLinkAlquilerConReferido_(codigo) {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('FORM_ALQUILER_ID');
  if (!formId) return '';
  try {
    const form = FormApp.openById(formId);
    const itemCodigo = form.getItems(FormApp.ItemType.TEXT)
      .find(it => it.getTitle() === 'Código de Embajador (no modificar)');
    if (!itemCodigo) return form.getPublishedUrl();
    const respuesta = form.createResponse();
    respuesta.withItemResponse(itemCodigo.asTextItem().createResponse(codigo));
    return respuesta.toPrefilledUrl();
  } catch (e) {
    return '';
  }
}

/**
 * Landing page exclusiva de cada embajador (?ref=CODIGO). Es lo más cercano
 * a "Landing Page exclusiva" que se puede construir gratis con Apps Script
 * — una página propia, con el nombre del embajador si existe, que lleva al
 * formulario de Comprador con el código ya prellenado.
 */
function construirLandingEmbajador_(codigo) {
  const filaEmbajador = buscarFilaEmbajadorPorCodigo_(codigo);
  let nombreEmbajador = '';
  let valido = filaEmbajador > -1;
  if (valido) {
    const obj = filaAObjeto_(SHEETS.EMBAJADORES, sh_(SHEETS.EMBAJADORES).getRange(filaEmbajador, 1, 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()[0]);
    nombreEmbajador = obj.Nombre || '';
    valido = obj.Estado === 'Activo';
  }

  const linkFormulario = valido ? construirLinkCompradorConReferido_(codigo) : '';
  const saludo = nombreEmbajador ? nombreEmbajador + ' te invitó a ' + APP_NAME : 'Te invitaron a ' + APP_NAME;

  const cuerpo = '<div style="text-align:center;padding-top:30px">' +
    '<h1 style="margin-bottom:6px">🏠 ' + APP_NAME + '</h1>' +
    '<p style="font-size:18px;color:#333">' + saludo + '</p>' +
    '<p style="color:#666">Te conectamos con un agente inmobiliario verificado en menos de 3 minutos.</p>' +
    (linkFormulario
      ? '<p style="margin-top:24px"><a href="' + linkFormulario + '" style="background:#1a73e8;color:white;padding:14px 28px;border-radius:8px;text-decoration:none;font-size:17px">Comenzar ahora →</a></p>'
      : '<p style="color:#990000">Este enlace no está disponible en este momento.</p>') +
    '</div>';

  const metaRefresh = linkFormulario ? '<meta http-equiv="refresh" content="4;url=' + linkFormulario + '">' : '';

  return '<!DOCTYPE html><html><head><base target="_top">' + metaRefresh +
    '<style>body{font-family:Arial,sans-serif;max-width:480px;margin:0 auto;padding:20px;background:#fafafa;color:#1a1a1a}' +
    'a{color:#1a73e8}</style></head><body>' + cuerpo + '</body></html>';
}

/** Muestra las instrucciones para desplegar el Portal de Resultados (menú) */
function mostrarInstruccionesPortal() {
  const urlActual = getConfig_('URL_PORTAL_RESULTADOS', '');
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:10px;line-height:1.5">' +
    '<h3>🌐 Activar el Portal de Resultados</h3>' +
    '<p><b>Estado actual:</b> ' + (urlActual ? '✅ Configurado (' + urlActual + ')' : '⚠️ Aún no configurado') + '</p>' +
    '<ol>' +
    '<li>En el editor de Apps Script, clic en <b>Implementar → Nueva implementación</b>.</li>' +
    '<li>Tipo: <b>Aplicación web</b>.</li>' +
    '<li>Ejecutar como: <b>Yo (tu cuenta)</b>.</li>' +
    '<li>Quién tiene acceso: <b>Cualquier usuario</b>.</li>' +
    '<li>Clic en <b>Implementar</b> y autoriza los permisos si te lo pide.</li>' +
    '<li>Copia la URL que te da Google (termina en <code>/exec</code>).</li>' +
    '<li>Pégala en la hoja <b>Configuración</b>, parámetro <b>URL_PORTAL_RESULTADOS</b>.</li>' +
    '</ol>' +
    '<p>Sin este paso, todo el sistema sigue funcionando normal — solo no aparece el botón "Ver mi análisis completo" en los correos.</p>' +
    '</div>'
  ).setWidth(480).setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Portal de Resultados');
}

/** Reenvía el link del Panel del Agente por si lo perdió (menú) */
function reenviarLinkPanelAgente() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Reenviar Panel de Agente', 'ID del Agente (ej. AGRD-000001):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const idAgente = resp.getResponseText().trim().toUpperCase();

  const fila = buscarFilaPorId_(SHEETS.AGENTES, idAgente);
  if (fila === -1) { alertaSegura_('', '⚠️ No se encontró el agente ' + idAgente + '.'); return; }

  const agentesSheet = sh_(SHEETS.AGENTES);
  let obj = filaAObjeto_(SHEETS.AGENTES, agentesSheet.getRange(fila, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);
  if (!obj.Codigo_Panel_Privado) {
    obj.Codigo_Panel_Privado = generarCodigoPrivadoLargo_();
    agentesSheet.getRange(fila, col_(SHEETS.AGENTES, 'Codigo_Panel_Privado')).setValue(obj.Codigo_Panel_Privado);
  }

  const link = construirLinkPanelAgente_(obj.Codigo_Panel_Privado);
  if (!link) { alertaSegura_('', '⚠️ El Portal de Resultados no está desplegado todavía (URL_PORTAL_RESULTADOS vacío en Configuración).'); return; }

  if (obj.Email) {
    enviarCorreo_(obj.Email, '📊 Tu Panel de Agente — ' + idAgente,
      '<div style="font-family:Arial"><h2>📊 Aquí está tu Panel de Agente</h2>' +
      '<p style="text-align:center"><a href="' + link + '" style="display:inline-block;background:#1a73e8;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:15px">📊 Ver mi Panel</a></p><p style="font-size:12px;color:#999;text-align:center">Guárdalo — es personal e intransferible.</p></div>');
  }
  mostrarPopupConWhatsApp_('📊 Panel de Agente reenviado', link, obj.Telefono,
    'Hola ' + obj.Nombre + ', aquí está el link a tu Panel de Agente: ' + link);
}

/** Reenvía el link del Portal de Resultados por si el comprador lo perdió (menú) */
function reenviarLinkPortalComprador() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Reenviar Portal de Resultados', 'ID del Comprador (ej. CVRD-000001):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const idComprador = resp.getResponseText().trim().toUpperCase();

  const fila = buscarFilaPorId_(SHEETS.COMPRADORES, idComprador);
  if (fila === -1) { alertaSegura_('', '⚠️ No se encontró el comprador ' + idComprador + '.'); return; }

  const compradoresSheet = sh_(SHEETS.COMPRADORES);
  let obj = filaAObjeto_(SHEETS.COMPRADORES, compradoresSheet.getRange(fila, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
  if (!obj.Codigo_Portal_Privado) {
    obj.Codigo_Portal_Privado = generarCodigoPrivadoLargo_();
    compradoresSheet.getRange(fila, col_(SHEETS.COMPRADORES, 'Codigo_Portal_Privado')).setValue(obj.Codigo_Portal_Privado);
  }

  const link = construirLinkPortalResultados_(obj.Codigo_Portal_Privado);
  if (!link) { alertaSegura_('', '⚠️ El Portal de Resultados no está desplegado todavía (URL_PORTAL_RESULTADOS vacío en Configuración).'); return; }

  if (obj.Email) {
    enviarCorreo_(obj.Email, '📊 Tu análisis — ' + idComprador,
      '<div style="font-family:Arial"><h2>📊 Aquí está tu análisis completo</h2>' +
      '<p style="text-align:center"><a href="' + link + '" style="display:inline-block;background:#1a73e8;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:15px">📊 Ver mi Análisis</a></p><p style="font-size:12px;color:#999;text-align:center">Guarda este link — es personal.</p></div>');
  }
  mostrarPopupConWhatsApp_('📊 Portal reenviado', link, obj.Telefono,
    'Hola ' + obj.Nombre + ', aquí está el link a tu análisis completo: ' + link);
}

/**
 * ============================================================================
 *  MODO API (JSON) — para el frontend de GitHub Pages
 * ============================================================================
 *  Este backend deja de "verse bonito" y solo entrega datos. El diseño,
 *  la interactividad y la marca viven en GitHub Pages (HTML/CSS/JS propio,
 *  gratis, con tu propio dominio) — evita el problema de Android que
 *  interceptaba links de script.google.com, porque el link que se comparte
 *  ahora es de github.io, no de Google.
 * ============================================================================
 */
function manejarSolicitudApi_(e) {
  const salida = (objeto) => ContentService.createTextOutput(JSON.stringify(objeto)).setMimeType(ContentService.MimeType.JSON);

  try {
    // Límite de volumen general: máximo 60 solicitudes por minuto a este
    // endpoint — evita que un script recorra miles de códigos al azar
    // buscando adivinar uno válido, sin afectar el uso normal (una persona
    // real nunca hace 60 solicitudes en un minuto).
    if (!verificarLimiteIntentos_('solicitudes_globales', 60, 1)) {
      return salida({ error: true, mensaje: 'Demasiadas solicitudes en poco tiempo. Espera un momento e intenta de nuevo.' });
    }

    const accion = (e.parameter.accion || '').trim();
    if (accion) {
      // Límite adicional por código: máximo 10 acciones de escritura por
      // minuto para el MISMO código — evita que alguien machaque un mismo
      // registro con cambios repetidos.
      const codigoParaLimite = (e.parameter.panel || e.parameter.portal || e.parameter.portalAlquiler || e.parameter.portalPropietario || e.parameter.embajador || e.parameter.panelAdmin || '').trim();
      if (codigoParaLimite && !verificarLimiteIntentos_('accion_' + codigoParaLimite, 10, 1)) {
        return salida({ error: true, mensaje: 'Demasiados intentos seguidos. Espera un minuto e intenta de nuevo.' });
      }
      return salida(ejecutarAccionMantenimiento_(e, accion));
    }

    const ref = (e.parameter.ref || '').toUpperCase().trim();
    if (ref) return salida(obtenerDatosEmbajadorJSON_(ref));

    const codigoPanelEmbajador = (e.parameter.embajador || '').trim();
    if (codigoPanelEmbajador) return salida(obtenerDatosPanelEmbajadorJSON_(codigoPanelEmbajador));

    const codigoPanel = (e.parameter.panel || '').trim();
    if (codigoPanel) {
      const idContacto = (e.parameter.contacto || '').toUpperCase().trim();
      return salida(obtenerDatosPanelAgenteJSON_(codigoPanel, idContacto));
    }

    const codigoPortal = (e.parameter.portal || '').trim();
    if (codigoPortal) return salida(obtenerDatosPortalCompradorJSON_(codigoPortal));

    const codigoPortalAlquiler = (e.parameter.portalAlquiler || '').trim();
    if (codigoPortalAlquiler) return salida(obtenerDatosPortalAlquilerJSON_(codigoPortalAlquiler));

    const codigoPortalPropietario = (e.parameter.portalPropietario || '').trim();
    if (codigoPortalPropietario) return salida(obtenerDatosPortalPropietarioJSON_(codigoPortalPropietario));

    const codigoPanelAdmin = (e.parameter.panelAdmin || '').trim();
    if (codigoPanelAdmin) return salida(obtenerDatosPanelAdminJSON_(codigoPanelAdmin));

    return salida({ error: true, mensaje: 'Falta un parámetro válido (portal, portalAlquiler, portalPropietario, panel, embajador, panelAdmin o ref).' });
  } catch (err) {
    log_('manejarSolicitudApi_', 'ERROR', err.message + ' | ' + err.stack);
    return salida({ error: true, mensaje: 'Ocurrió un error procesando la solicitud.' });
  }
}

/** Datos JSON del Portal del Comprador */
function obtenerDatosPortalCompradorJSON_(codigoPortal) {
  const fila = buscarFilaPorCodigoPortal_(codigoPortal);
  if (fila === -1) return { error: true, mensaje: 'No encontramos ese código. Verifica el enlace.' };

  const obj = filaAObjeto_(SHEETS.COMPRADORES, sh_(SHEETS.COMPRADORES).getRange(fila, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
  const { faltantes } = calcularIndicePreparacion_(obj);
  const alternativas = obtenerProvinciasAlternativasSimple_(obj.Presupuesto_Max, obj.Provincia);
  const analisis = obtenerAnalisisPersonalizadoConCache_(obj, faltantes, alternativas, fila);
  const gastos = calcularGastosAdicionalesCompra_(obj);
  const { sugMin, sugMax } = evaluarRealidadPresupuesto_(obj);
  const resumenCompatibilidad = calcularResumenCompatibilidad_(obj);

  // Propiedades de propietarios en modo "Yo solo" compatibles con lo que
  // busca este comprador — mismos datos que ya se usan para el match
  // (precio, zona, tipo), sin fotos (el sistema no las maneja).
  const conexionesDelComprador = {};
  try {
    const sheetConex = sh_(SHEETS.CONEXIONES_DIRECTAS);
    if (sheetConex.getLastRow() > 1) {
      sheetConex.getRange(2, 1, sheetConex.getLastRow() - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues().forEach(f => {
        const c = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f);
        if (c.ID_Comprador === obj.ID_Comprador) conexionesDelComprador[c.ID_Propiedad] = { estado: c.Estado, idConexion: c.ID_Conexion };
      });
    }
  } catch (e) { /* hoja aún no existe */ }

  // Cierre esperando que el comprador confirme si el monto reportado por
  // el agente es correcto — se muestra directamente aquí en el Portal
  // (con botones Sí/No), en vez de depender únicamente del enlace del
  // correo, para que sea más fácil e inmediato para ambas partes.
  let montoCierrePendiente = null;
  try {
    const sheetCierres = sh_(SHEETS.CIERRES);
    if (sheetCierres.getLastRow() > 1) {
      const filaPendiente = sheetCierres.getRange(2, 1, sheetCierres.getLastRow() - 1, HEADERS[SHEETS.CIERRES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CIERRES, f))
        .find(c => c.ID_Comprador === obj.ID_Comprador && c.Monto_Confirmado_Comprador === 'Pendiente');
      if (filaPendiente) {
        montoCierrePendiente = { idCierre: filaPendiente.ID_Cierre, montoVenta: Number(filaPendiente.Monto_Venta) || 0 };
      }
    }
  } catch (e) { /* hoja aún no existe */ }

  const propiedadesCompatibles = buscarPropiedadesCompatiblesConComprador_(obj, false)
    .filter(p => p.Modo_Venta === 'Yo solo (sin agente)')
    .map(p => {
      const conexion = conexionesDelComprador[p.ID_Propiedad] || null;
      return {
        id: p.ID_Propiedad, tipo: p.Tipo_Inmueble, provincia: p.Provincia, municipio: p.Municipio,
        precio: Number(p.Precio) || 0, habitaciones: p.Habitaciones, banos: p.Banos,
        descripcion: p.Descripcion || '',
        conexion: conexion,
        // Solo se calcula (y se muestra) el link de WhatsApp del propietario
        // cuando AMBOS ya aceptaron conectar — antes de eso, nunca se
        // revela el contacto, manteniendo el mismo principio de
        // consentimiento mutuo del resto del sistema.
        linkWhatsAppPropietario: (conexion && conexion.estado === 'Conectados' && p.Telefono)
          ? linkWhatsApp_(p.Telefono, 'Hola ' + p.Nombre_Propietario + ', soy ' + obj.Nombre + ', nos conectamos por PropMatchRD sobre tu propiedad.')
          : ''
      };
    });

  // Datos del agente asignado — para que el comprador pueda escribirle
  // directo por WhatsApp desde su Portal, sin depender de esperar a que el
  // agente lo contacte primero. Se envía un perfil más completo (no solo
  // el nombre) para que el comprador conozca con quién está trabajando —
  // genera más confianza que solo un nombre suelto.
  let agenteNombre = '', agenteLinkWhatsApp = '', agentePerfil = null;
  if (obj.Agente_Asignado_ID) {
    const filaAgente = buscarFilaPorId_(SHEETS.AGENTES, obj.Agente_Asignado_ID);
    if (filaAgente > -1) {
      const agenteObj = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);
      agenteNombre = agenteObj.Nombre + ' ' + agenteObj.Apellido;
      agenteLinkWhatsApp = linkWhatsApp_(agenteObj.Telefono, 'Hola ' + agenteObj.Nombre + ', soy ' + obj.Nombre + ', tengo una consulta sobre mi búsqueda de propiedad.');
      agentePerfil = {
        nombre: agenteNombre,
        empresa: agenteObj.Empresa || '',
        especialidad: agenteObj.Especialidad || '',
        anosExperiencia: agenteObj.Anos_Experiencia || '',
        cierresConfirmados: Number(agenteObj.Cierres_Confirmados) || 0,
        rating: Number(agenteObj.Rating) || 0,
        satisfaccionPct: Number(agenteObj.Satisfaccion_Clientes_Pct) || 0,
        verificado: agenteObj.Verificado === 'Sí',
        distintivo: agenteObj.Distintivo || ''
      };
    }
  }

  // Historial de visitas — con el nombre del agente de CADA visita (no
  // solo el asignado actual, por si hubo una reasignación en el camino).
  const historialVisitas = obtenerHistorialVisitasConAgente_(SHEETS.VISITAS, 'ID_Comprador', obj.ID_Comprador, 'Resultado', 'Comprador_Le_Gusto');

  return {
    error: false,
    idComprador: obj.ID_Comprador,
    nombre: obj.Nombre,
    clasificacion: obj.Clasificacion,
    estado: obj.Estado,
    estadoExplicado: ESTADO_COMPRADOR_EXPLICADO[obj.Estado] || obj.Estado,
    telefono: obj.Telefono || '',
    email: obj.Email || '',
    agenteNombre: agenteNombre,
    agenteLinkWhatsApp: agenteLinkWhatsApp,
    agentePerfil: agentePerfil,
    historialVisitas: historialVisitas,
    conexionesDirectasPendientes: obtenerConexionesDirectasPendientes_(obj.ID_Comprador),
    indicePreparacionPct: Number(obj.Indice_Preparacion_Pct) || 0,
    probabilidadFinanciamientoPct: Number(obj.Probabilidad_Financiamiento_Pct) || 0,
    analisisPersonalizado: analisis,
    etapaRutaCompra: obj.Etapa_Ruta_Compra || 'Evaluación',
    etapasRutaCompra: ETAPAS_RUTA_COMPRA,
    faltantes: faltantes,
    alertaPresupuesto: obj.Alerta_Presupuesto || '',
    graficoPresupuesto: {
      presupuestoMax: Number(obj.Presupuesto_Max) || 0,
      presupuestoMin: Number(obj.Presupuesto_Min) || 0,
      mercadoMin: sugMin || 0,
      mercadoMax: sugMax || 0
    },
    resumenCompatibilidad: resumenCompatibilidad,
    confirmarCierrePendiente: obj.Cierre_Confirmado_Comprador === 'Pendiente',
    urlFormularioEmbajador: PropertiesService.getScriptProperties().getProperty('FORM_EMBAJADOR_URL') || '',
    montoCierrePendiente: montoCierrePendiente,
    propiedadesCompatibles: propiedadesCompatibles,
    certificadoUrl: obj.Certificado_URL || '',
    documentos: {
      cedula: obj.Doc_Cedula === 'Sí', cartaTrabajo: obj.Doc_Carta_Trabajo === 'Sí', estadosCuenta: obj.Doc_Estados_Cuenta === 'Sí'
    },
    calificacionVisitaPendiente: obtenerCalificacionVisitaPendiente_(obj.ID_Comprador),
    gastos: {
      inicial: gastos.inicial, transferencia: gastos.transferencia,
      seguroMin: gastos.seguroMin, seguroMax: gastos.seguroMax,
      tasacionMin: gastos.tasacionMin, tasacionMax: gastos.tasacionMax,
      cierreMin: gastos.cierreMin, cierreMax: gastos.cierreMax,
      aplicaIPI: gastos.aplicaIPI, ipiAnual: gastos.ipiAnual,
      totalMin: gastos.totalMin, totalMax: gastos.totalMax
    }
  };
}

/**
 * Revisa si la última visita del comprador todavía tiene alguna de las 3
 * preguntas de reputación sin responder — para mostrar la mini-encuesta en
 * el Portal solo cuando de verdad hace falta, no siempre.
 */
function obtenerCalificacionVisitaPendiente_(idComprador) {
  const sheet = sh_(SHEETS.VISITAS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.VISITAS].length).getValues().map(f => filaAObjeto_(SHEETS.VISITAS, f));
  const visitasComprador = data.filter(v => v.ID_Comprador === idComprador && v.Resultado);
  if (!visitasComprador.length) return false;
  const ultima = visitasComprador[visitasComprador.length - 1];
  return !ultima.Agente_Puntual || !ultima.Agente_Explico_Bien || !ultima.Recomendaria_Agente || !ultima.Comprador_Le_Gusto;
}

/** Datos JSON del Panel del Agente (incluye la acción de marcar contacto si viene idContacto) */
function obtenerDatosPanelAgenteJSON_(codigoPanel, idContacto) {
  const filaAgente = buscarFilaAgentePorCodigoPanel_(codigoPanel);
  if (filaAgente === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo (revisa que no se haya cortado al copiarlo), o pide que te reenvíen tu Panel desde el menú del sistema.' };

  const agente = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);

  let mensajeContacto = '';
  if (idContacto) {
    mensajeContacto = marcarPrimerContactoDesdeAgente_(idContacto, agente.ID_Agente);
  }

  // Historial de visitas — se construye UNA sola vez leyendo cada hoja
  // completa (en vez de una consulta por comprador/solicitud), agrupado
  // por ID para consultarlo al vuelo al armar cada tarjeta.
  const construirMapaHistorial_ = (sheetKey, campoId) => {
    const mapa = {};
    let sheet;
    try {
      sheet = sh_(sheetKey);
    } catch (e) {
      // Si la hoja aún no existe (ej. no se ha ejecutado "Parte 1A" tras
      // esta actualización), el Panel completo no debe romperse por esto
      // — simplemente no hay historial que mostrar todavía.
      return mapa;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return mapa;
    sheet.getRange(2, 1, lastRow - 1, HEADERS[sheetKey].length).getValues().forEach(f => {
      const obj = filaAObjeto_(sheetKey, f);
      const id = obj[campoId];
      if (!mapa[id]) mapa[id] = [];
      mapa[id].push(obj);
    });
    return mapa;
  };
  const historialVisitasCompra = construirMapaHistorial_(SHEETS.VISITAS, 'ID_Comprador');
  const historialVisitasAlquiler = construirMapaHistorial_(SHEETS.VISITAS_ALQUILER, 'ID_Solicitud');

  // totalVisitasRealizadas se calcula con los mismos datos que acabamos de
  // leer arriba — evita releer por completo las hojas de Visitas y
  // Visitas de Alquiler otra vez (que era lo que hacía
  // contarVisitasDeAgente_ por su cuenta).
  let totalVisitasRealizadas = 0;
  Object.values(historialVisitasCompra).forEach(visitas => {
    totalVisitasRealizadas += visitas.filter(v => v.ID_Agente === agente.ID_Agente).length;
  });
  Object.values(historialVisitasAlquiler).forEach(visitas => {
    totalVisitasRealizadas += visitas.filter(v => v.ID_Agente === agente.ID_Agente).length;
  });

  const compradoresSheet = sh_(SHEETS.COMPRADORES);
  const lastRowComp = compradoresSheet.getLastRow();
  const misCompradores = lastRowComp > 1
    ? compradoresSheet.getRange(2, 1, lastRowComp - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
        .filter(c => c.Agente_Asignado_ID === agente.ID_Agente && !['Cerrado', 'Perdido', 'Duplicado'].includes(c.Estado))
    : [];

  const cierresSheet = sh_(SHEETS.CIERRES);
  const lastRowCierres = cierresSheet.getLastRow();
  const comisionesPendientes = lastRowCierres > 1
    ? cierresSheet.getRange(2, 1, lastRowCierres - 1, HEADERS[SHEETS.CIERRES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CIERRES, f))
        .filter(c => c.ID_Agente === agente.ID_Agente && c.Estado_Pago_Comision !== 'Pagado')
    : [];

  const asignacionesSheet = sh_(SHEETS.ASIGNACIONES);
  const lastRowAsig = asignacionesSheet.getLastRow();
  const contactosPendientes = {};
  const idsCompitiendo = [];
  let totalContactados = 0;
  if (lastRowAsig > 1) {
    asignacionesSheet.getRange(2, 1, lastRowAsig - 1, HEADERS[SHEETS.ASIGNACIONES].length).getValues()
      .map(f => filaAObjeto_(SHEETS.ASIGNACIONES, f))
      .filter(a => a.ID_Agente === agente.ID_Agente)
      .forEach(a => {
        contactosPendientes[a.ID_Comprador] = !a.Fecha_Primer_Contacto;
        if (a.Fecha_Primer_Contacto) totalContactados++;
        if (a.Estado === 'Compitiendo') idsCompitiendo.push(a.ID_Comprador);
      });
  }

  // "Tu próxima mejor acción" — en vez de que el agente tenga que decidir
  // por su cuenta a quién contactar primero mirando toda la lista, se le
  // sugiere directamente: prioriza a quien todavía no ha contactado (son
  // los más sensibles al tiempo — cada día sin contacto es una
  // oportunidad de que se enfríe o lo asignen a otro agente en
  // competencia), y entre esos, al de mayor PropScore.
  let proximaMejorAccion = null;
  const candidatosSinContactar = misCompradores.filter(c => contactosPendientes[c.ID_Comprador]);
  if (candidatosSinContactar.length) {
    candidatosSinContactar.sort((a, b) => (Number(b.PropScore) || 0) - (Number(a.PropScore) || 0));
    const top = candidatosSinContactar[0];
    proximaMejorAccion = {
      idComprador: top.ID_Comprador, nombre: top.Nombre,
      razon: candidatosSinContactar.length > 1
        ? 'Todavía no lo contactas, y tiene el PropScore más alto entre los ' + candidatosSinContactar.length + ' pendientes de contactar.'
        : 'Todavía no lo has contactado — cada día sin contacto es una oportunidad que se enfría.'
    };
  } else if (misCompradores.length) {
    const activos = misCompradores.filter(c => c.Estado !== 'Cerrado' && c.Estado !== 'Perdido');
    if (activos.length) {
      activos.sort((a, b) => (Number(b.PropScore) || 0) - (Number(a.PropScore) || 0));
      const top = activos[0];
      proximaMejorAccion = { idComprador: top.ID_Comprador, nombre: top.Nombre, razon: 'Tiene el PropScore más alto de tu cartera — el más listo para avanzar.' };
    }
  }
  // Comisión ESTIMADA del agente por cada comprador/alquiler — antes de
  // cerrar, para que vea el potencial de cada uno, no solo el IICD/etapa.
  // Se calcula igual que el Motor 1 real, mostrando el NETO que le
  // quedaría tras la parte de la plataforma — la misma cifra que vería en
  // su correo si cierra. Se define ANTES de construir las listas de abajo,
  // porque todas la necesitan.
  const pctComisionAgenteDefault = Number(getConfig_('COMISION_AGENTE_PCT_DEFAULT', 5));
  const pctPlataformaComision = Number(getConfig_('PCT_PLATAFORMA_COMISION_AGENTE', 20));
  const calcularComisionNetaVenta_ = (presupuestoMin, presupuestoMax) => {
    const bruteMin = Number(presupuestoMin || 0) * pctComisionAgenteDefault / 100;
    const bruteMax = Number(presupuestoMax || 0) * pctComisionAgenteDefault / 100;
    return {
      min: Math.round(bruteMin * (1 - pctPlataformaComision / 100)),
      max: Math.round(bruteMax * (1 - pctPlataformaComision / 100))
    };
  };
  const pctComisionAlquilerAgente = Number(getConfig_('PCT_COMISION_ALQUILER_AGENTE', 100));
  const calcularComisionNetaAlquiler_ = (presupuestoMin, presupuestoMax) => ({
    min: Math.round(Number(presupuestoMin || 0) * pctComisionAlquilerAgente / 100),
    max: Math.round(Number(presupuestoMax || 0) * pctComisionAlquilerAgente / 100)
  });

  // Compradores en COMPETENCIA paralela (Plata/Bronce) — antes esto era
  // invisible en el Panel; el agente solo se enteraba por correo y nunca
  // volvía a verlo hasta ganar. Ahora aparece en su propia sección,
  // dejando claro que todavía no es exclusivo suyo.
  const compradoresCompitiendo = idsCompitiendo.length
    ? compradoresSheet.getRange(2, 1, lastRowComp - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
        .filter(c => idsCompitiendo.includes(c.ID_Comprador))
        .map(c => {
          const est = calcularComisionNetaVenta_(c.Presupuesto_Min, c.Presupuesto_Max);
          return { id: c.ID_Comprador, nombre: c.Nombre + ' ' + c.Apellido, zona: c.Provincia + ' - ' + c.Municipio, clasificacion: c.Clasificacion, formaPago: c.Forma_Pago || '', comisionEstMin: est.min, comisionEstMax: est.max, linkWhatsApp: linkWhatsApp_(c.Telefono, 'Hola ' + c.Nombre + ', soy ' + agente.Nombre + ' de ' + APP_NAME + '. Vi tu solicitud y me encantaría ayudarte.') };
        })
    : [];

  // Módulo de Alquiler: solicitudes asignadas exclusivamente a este agente,
  // y las que están en competencia paralela (mismo patrón que Compradores).
  const solicitudesAlquilerSheet = sh_(SHEETS.SOLICITUDES_ALQUILER);
  const lastRowAlquiler = solicitudesAlquilerSheet.getLastRow();
  const todasSolicitudesAlquiler = lastRowAlquiler > 1
    ? solicitudesAlquilerSheet.getRange(2, 1, lastRowAlquiler - 1, HEADERS[SHEETS.SOLICITUDES_ALQUILER].length).getValues()
        .map(f => filaAObjeto_(SHEETS.SOLICITUDES_ALQUILER, f))
    : [];
  const misAlquileres = todasSolicitudesAlquiler
    .filter(a => a.Agente_Asignado_ID === agente.ID_Agente && a.Estado !== 'Cerrado')
    .map(a => {
      const est = calcularComisionNetaAlquiler_(a.Presupuesto_Min_Mensual, a.Presupuesto_Max_Mensual);
      return {
        id: a.ID_Solicitud, nombre: a.Nombre + ' ' + a.Apellido, zona: a.Provincia + ' - ' + a.Municipio,
        presupuesto: 'RD$' + Number(a.Presupuesto_Min_Mensual).toLocaleString() + ' - RD$' + Number(a.Presupuesto_Max_Mensual).toLocaleString(),
        clasificacion: a.Clasificacion, comisionEstMin: est.min, comisionEstMax: est.max,
        fechaMudanzaDeseada: a.Fecha_Mudanza_Deseada || '',
        linkWhatsApp: linkWhatsApp_(a.Telefono, 'Hola ' + a.Nombre + ', soy ' + agente.Nombre + ' de ' + APP_NAME + '. Vi tu solicitud de alquiler y me encantaría ayudarte.'),
        historialVisitas: (historialVisitasAlquiler[a.ID_Solicitud] || []).map(v => ({
          numero: v.Numero_Visita, fecha: v.Fecha_Registro, resultado: 'Visita registrada', leGusto: v.Inquilino_Le_Gusto || '',
          confirmoVisita: v.Inquilino_Confirmo_Visita || '', agentePuntual: v.Agente_Puntual || '',
          agenteExplicoBien: v.Agente_Explico_Bien || '', recomendaria: v.Recomendaria_Agente || ''
        }))
      };
    });
  const alquileresCompitiendo = idsCompitiendo.length
    ? todasSolicitudesAlquiler
        .filter(a => idsCompitiendo.includes(a.ID_Solicitud))
        .map(a => {
          const est = calcularComisionNetaAlquiler_(a.Presupuesto_Min_Mensual, a.Presupuesto_Max_Mensual);
          return { id: a.ID_Solicitud, nombre: a.Nombre + ' ' + a.Apellido, zona: a.Provincia + ' - ' + a.Municipio, clasificacion: a.Clasificacion, comisionEstMin: est.min, comisionEstMax: est.max, linkWhatsApp: linkWhatsApp_(a.Telefono, 'Hola ' + a.Nombre + ', soy ' + agente.Nombre + ' de ' + APP_NAME + '. Vi tu solicitud de alquiler y me encantaría ayudarte.') };
        })
    : [];

  // Comisiones CONFIRMADAS (ya cerradas) — con fecha y estado real de pago,
  // para que el agente tenga un panorama financiero completo, no solo lo
  // que está en curso. Combina venta y alquiler en una sola lista.
  const misCierresVenta = lastRowCierres > 1
    ? cierresSheet.getRange(2, 1, lastRowCierres - 1, HEADERS[SHEETS.CIERRES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CIERRES, f))
        .filter(c => c.ID_Agente === agente.ID_Agente)
        .map(c => ({
          tipo: 'Compra', id: c.ID_Comprador, monto: Number(c.Comision_Plataforma_RD) || 0,
          fecha: c.Fecha_Cierre, estado: c.Estado_Pago_Comision || 'Pendiente',
          desglose: {
            montoVenta: Number(c.Monto_Venta) || 0, pctComisionAgente: Number(c.Comision_Agente_Pct) || 0,
            comisionAgenteRD: Number(c.Comision_Agente_RD) || 0, pctPlataforma: Number(c.Comision_Plataforma_Pct) || 0
          }
        }))
    : [];
  const cierresAlquilerSheet = sh_(SHEETS.CIERRES_ALQUILER);
  const lastRowCierresAlq = cierresAlquilerSheet.getLastRow();
  const misCierresAlquiler = lastRowCierresAlq > 1
    ? cierresAlquilerSheet.getRange(2, 1, lastRowCierresAlq - 1, HEADERS[SHEETS.CIERRES_ALQUILER].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CIERRES_ALQUILER, f))
        .filter(c => c.ID_Agente === agente.ID_Agente)
        .map(c => ({
          tipo: 'Alquiler', id: c.ID_Solicitud, monto: Number(c.Comision_Plataforma_RD) || 0,
          fecha: c.Fecha_Cierre, estado: c.Estado_Pago_Comision || 'Pendiente',
          desglose: {
            montoVenta: Number(c.Monto_Alquiler_Mensual) || 0, pctComisionAgente: Number(c.Comision_Agente_Pct) || 0,
            comisionAgenteRD: Number(c.Comision_Agente_RD) || 0, pctPlataforma: Number(c.Comision_Plataforma_Pct) || 0
          }
        }))
    : [];
  const misComisiones = misCierresVenta.concat(misCierresAlquiler)
    .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

  return {
    error: false,
    nombre: agente.Nombre,
    tipoRegistro: agente.Tipo_Registro || 'Individual',
    nombreRepresentante: agente.Nombre_Representante || '',
    nivel: agente.Distintivo || '🥉 Bronze',
    telefono: agente.Telefono || '',
    email: agente.Email || '',
    disponibilidad: agente.Disponibilidad || 'Sí',
    atiendeAlquileres: agente.Atiende_Alquileres || 'No',
    premium: {
      activo: esPremiumActivo_(agente),
      fechaVencimiento: agente.Fecha_Vencimiento_Premium || '',
      pagoReportado: agente.Pago_Premium_Reportado === 'Sí',
      cuota: Number(getConfig_('CUOTA_PREMIUM_AGENTE_RD', 500)),
      pctPlataformaPremium: Number(getConfig_('PCT_PLATAFORMA_COMISION_AGENTE_PREMIUM', 12)),
      pctPlataformaNormal: Number(getConfig_('PCT_PLATAFORMA_COMISION_AGENTE', 20)),
      bonoMatching: Number(getConfig_('BONO_MATCHING_PREMIUM_PCT', 40))
    },
    mensajeContacto: mensajeContacto,
    kpis: {
      compradoresActivos: misCompradores.length,
      alquileresActivos: misAlquileres.length,
      cierresConfirmados: Number(agente.Cierres_Confirmados) || 0,
      tasaConversionPct: Number(agente.Tasa_Conversion) || 0,
      tiempoRespuestaHoras: agente.Tiempo_Respuesta_Promedio_Horas || null,
      totalContactados: totalContactados,
      totalVisitasRealizadas: totalVisitasRealizadas
    },
    ranking: obtenerPosicionRanking_(agente.ID_Agente),
    proximaMejorAccion: proximaMejorAccion,
    comisionesPendientes: comisionesPendientes.length,
    misComisiones: misComisiones,
    cuentaBancariaPlataforma: obtenerCuentaPlataformaFormateada_(),
    compradoresCompitiendo: compradoresCompitiendo,
    alquileresActivos: misAlquileres,
    alquileresCompitiendo: alquileresCompitiendo,
    compradores: misCompradores
      .sort((a, b) => (Number(b.IICD_Score) || 0) - (Number(a.IICD_Score) || 0)) // prioridad más alta primero
      .map(c => {
        const comisionEst = calcularComisionNetaVenta_(c.Presupuesto_Min, c.Presupuesto_Max);
        const { sugMin, sugMax } = evaluarRealidadPresupuesto_(c);
        const resumenCompatibilidad = calcularResumenCompatibilidad_(c);
        return {
          id: c.ID_Comprador, nombre: c.Nombre + ' ' + c.Apellido,
          zona: c.Provincia + ' - ' + c.Municipio,
          etapaRutaCompra: c.Etapa_Ruta_Compra || '—', ipcEtiqueta: c.IPC_Etiqueta || '—',
          iicdScore: Number(c.IICD_Score) || 0, iicdEtiqueta: c.IICD_Etiqueta || '—', propScore: Number(c.PropScore) || 0,
          contactoPendiente: !!contactosPendientes[c.ID_Comprador],
          formaPago: c.Forma_Pago || '',
          viveEnElExterior: c.Vive_En_El_Exterior === 'Sí',
          comisionEstMin: comisionEst.min, comisionEstMax: comisionEst.max,
          graficoPresupuesto: {
            presupuestoMax: Number(c.Presupuesto_Max) || 0, presupuestoMin: Number(c.Presupuesto_Min) || 0,
            mercadoMin: sugMin || 0, mercadoMax: sugMax || 0
          },
          resumenCompatibilidad: resumenCompatibilidad,
          linkWhatsApp: linkWhatsApp_(c.Telefono, 'Hola ' + c.Nombre + ', soy ' + agente.Nombre + ' de ' + APP_NAME + '. ¿Cómo va tu búsqueda de propiedad?'),
          historialVisitas: (historialVisitasCompra[c.ID_Comprador] || []).map(v => ({
            numero: v.Numero_Visita, fecha: v.Fecha_Registro_Resultado || v.Fecha_Programada,
            resultado: v.Resultado || '—', leGusto: v.Comprador_Le_Gusto || '',
            confirmoVisita: v.Comprador_Confirmo_Visita || '', agentePuntual: v.Agente_Puntual || '',
            agenteExplicoBien: v.Agente_Explico_Bien || '', recomendaria: v.Recomendaria_Agente || ''
          }))
        };
      })
  };
}

/** Datos JSON de la landing del Embajador */
function obtenerDatosEmbajadorJSON_(codigo) {
  const filaEmbajador = buscarFilaEmbajadorPorCodigo_(codigo);
  if (filaEmbajador === -1) return { error: true, mensaje: 'Código no válido.' };

  const obj = filaAObjeto_(SHEETS.EMBAJADORES, sh_(SHEETS.EMBAJADORES).getRange(filaEmbajador, 1, 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()[0]);
  if (obj.Estado !== 'Activo') return { error: true, mensaje: 'Este enlace no está disponible en este momento.' };

  const linkFormularioComprador = construirLinkCompradorConReferido_(codigo);
  const linkFormularioAlquiler = construirLinkAlquilerConReferido_(codigo);
  return {
    error: false,
    nombreEmbajador: obj.Nombre || '',
    linkFormularioComprador: linkFormularioComprador,
    linkFormularioAlquiler: linkFormularioAlquiler
  };
}

/** Busca la fila (1-based) de un embajador por su código PRIVADO de panel (no el código público de referido) */
function buscarFilaEmbajadorPorCodigoPanel_(codigo) {
  const sheet = sh_(SHEETS.EMBAJADORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const codigoBuscado = String(codigo).trim();
  const codigos = sheet.getRange(2, col_(SHEETS.EMBAJADORES, 'Codigo_Panel_Privado'), lastRow - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) return i + 2;
  }
  return -1;
}

/** Construye el link privado al Panel del Embajador (distinto del link público de referido) */
function construirLinkPanelEmbajador_(codigoPanel) {
  if (!codigoPanel) return '';
  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  if (urlGithub) return urlGithub.replace(/\/$/, '') + '/panel-embajador.html?codigo=' + encodeURIComponent(codigoPanel);
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  if (!urlBase) return '';
  const separador = urlBase.includes('?') ? '&' : '?';
  return urlBase + separador + 'embajador=' + encodeURIComponent(codigoPanel);
}

/**
 * Datos JSON del Panel PRIVADO del Embajador — su propio embudo (referidos,
 * validados, cerrados, comisión acumulada) y una línea de tiempo de su
 * progreso. Antes, un embajador solo recibía correos puntuales sin nunca
 * ver su propio embudo completo en un solo lugar — el mismo hueco que ya
 * habíamos corregido para agentes y compradores.
 */
function obtenerDatosPanelEmbajadorJSON_(codigoPanel) {
  const fila = buscarFilaEmbajadorPorCodigoPanel_(codigoPanel);
  if (fila === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo, o pide que te reenvíen tu Panel desde el menú del sistema.' };

  const embajador = filaAObjeto_(SHEETS.EMBAJADORES, sh_(SHEETS.EMBAJADORES).getRange(fila, 1, 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()[0]);

  const totalReferidos = Number(embajador.Total_Referidos) || 0;
  const validados = Number(embajador.Referidos_Validados) || 0;
  const cerrados = Number(embajador.Referidos_Cerrados) || 0;
  const comisionAcumulada = Number(embajador.Comision_Acumulada_RD) || 0;
  const score = Number(embajador.Embassador_Score_Pct) || 0;

  // Transparencia del Motor 2: mismos valores que se usan en el cálculo
  // real al momento del cierre — nunca desincronizado con lo que
  // efectivamente se le paga.
  const pctBasePool = Number(getConfig_('PCT_BASE_COMISION_EMBAJADORES', 1));
  const pctEmbajadorBase = Number(getConfig_('PCT_EMBAJADOR_DEL_POOL', 80));
  const scoreMinimoBono = Number(getConfig_('SCORE_MINIMO_BONO_EMBAJADOR', 80));
  const pctBonoAltoDesempeno = Number(getConfig_('PCT_BONO_EMBAJADOR_ALTO_DESEMPENO', 10));
  const calificaParaBono = score >= scoreMinimoBono;
  const pctEmbajadorEfectivo = calificaParaBono ? Math.min(100, pctEmbajadorBase + pctBonoAltoDesempeno) : pctEmbajadorBase;

  // Comisión reservada (Ecosistema de Embajadores): lo que ya está "en
  // camino" pero aún no se ha cerrado ni pagado — el embajador antes solo
  // veía lo ya cobrado, sin saber cuánto tenía pendiente en proceso.
  const comisionesSheet = sh_(SHEETS.COMISIONES_REFERIDOS);
  const lastRowComisiones = comisionesSheet.getLastRow();
  let comisionReservadaMin = 0, comisionReservadaMax = 0;
  if (lastRowComisiones > 1) {
    comisionesSheet.getRange(2, 1, lastRowComisiones - 1, HEADERS[SHEETS.COMISIONES_REFERIDOS].length).getValues()
      .map(f => filaAObjeto_(SHEETS.COMISIONES_REFERIDOS, f))
      .filter(c => c.Codigo_Referido === embajador.Codigo_Referido && c.Estado_Comision === 'Reservada')
      .forEach(c => {
        comisionReservadaMin += Number(c.Comision_Estimada_Min_RD) || 0;
        comisionReservadaMax += Number(c.Comision_Estimada_Max_RD) || 0;
      });
  }

  // Línea de tiempo: cada etapa se marca "hecha" según los contadores reales
  const pasos = [
    { titulo: 'Código generado', hecho: true },
    { titulo: 'Primer referido registrado', hecho: totalReferidos >= 1 },
    { titulo: 'Primera validación (agente asignado)', hecho: validados >= 1 },
    { titulo: 'Primera comisión ganada', hecho: cerrados >= 1 }
  ];

  // Listado de referidos — anonimizado a propósito (nombre + inicial del
  // apellido, sin teléfono/correo/presupuesto) para darle trazabilidad al
  // embajador sin exponer datos que le permitan contactar directamente al
  // comprador y saltarse al agente asignado.
  const compradoresSheetRef = sh_(SHEETS.COMPRADORES);
  const lastRowCompRef = compradoresSheetRef.getLastRow();

  // Mapa de comisiones por ID (comprador o solicitud de alquiler) — para
  // mostrar cuánto se reservó de CADA referido individualmente, no solo el
  // total acumulado de todos juntos.
  const comisionesSheetRef = sh_(SHEETS.COMISIONES_REFERIDOS);
  const lastRowComisionesRef = comisionesSheetRef.getLastRow();
  const comisionPorId = {};
  if (lastRowComisionesRef > 1) {
    comisionesSheetRef.getRange(2, 1, lastRowComisionesRef - 1, HEADERS[SHEETS.COMISIONES_REFERIDOS].length).getValues()
      .map(f => filaAObjeto_(SHEETS.COMISIONES_REFERIDOS, f))
      .filter(c => c.Codigo_Referido === embajador.Codigo_Referido)
      .forEach(c => { comisionPorId[c.ID_Comprador] = c; });
  }

  const listaReferidosCompra = lastRowCompRef > 1
    ? compradoresSheetRef.getRange(2, 1, lastRowCompRef - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
        .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
        .filter(c => c.Codigo_Referido === embajador.Codigo_Referido)
        .map(c => {
          const com = comisionPorId[c.ID_Comprador];
          return {
            nombre: c.Nombre + ' ' + (c.Apellido ? c.Apellido.charAt(0) + '.' : ''),
            tipo: 'Compra',
            estado: c.Estado === 'Cerrado' ? 'Cerrado ✅' : (c.Agente_Asignado_ID ? 'Validado (agente asignado)' : 'Registrado'),
            fecha: c.Fecha_Registro,
            comisionPagada: com ? com.Estado_Comision === 'Pagada' : false,
            comisionMin: com ? Number(com.Comision_Estimada_Min_RD) || 0 : null,
            comisionMax: com ? Number(com.Comision_Estimada_Max_RD) || 0 : null,
            comisionExacta: com && com.Comision_Referidor_RD ? Number(com.Comision_Referidor_RD) : null
          };
        })
    : [];

  // Módulo de Alquiler: mismos referidos, misma anonimización — el
  // embajador ve un solo listado combinado, con etiqueta de tipo.
  const solicitudesAlquilerSheetRef = sh_(SHEETS.SOLICITUDES_ALQUILER);
  const lastRowAlqRef = solicitudesAlquilerSheetRef.getLastRow();
  const listaReferidosAlquiler = lastRowAlqRef > 1
    ? solicitudesAlquilerSheetRef.getRange(2, 1, lastRowAlqRef - 1, HEADERS[SHEETS.SOLICITUDES_ALQUILER].length).getValues()
        .map(f => filaAObjeto_(SHEETS.SOLICITUDES_ALQUILER, f))
        .filter(a => a.Codigo_Referido === embajador.Codigo_Referido)
        .map(a => {
          const com = comisionPorId[a.ID_Solicitud];
          return {
            nombre: a.Nombre + ' ' + (a.Apellido ? a.Apellido.charAt(0) + '.' : ''),
            tipo: 'Alquiler',
            estado: a.Estado === 'Cerrado' ? 'Cerrado ✅' : (a.Agente_Asignado_ID ? 'Validado (agente asignado)' : 'Registrado'),
            fecha: a.Fecha_Registro,
            comisionPagada: com ? com.Estado_Comision === 'Pagada' : false,
            comisionMin: com ? Number(com.Comision_Estimada_Min_RD) || 0 : null,
            comisionMax: com ? Number(com.Comision_Estimada_Max_RD) || 0 : null,
            comisionExacta: com && com.Comision_Referidor_RD ? Number(com.Comision_Referidor_RD) : null
          };
        })
    : [];

  const listaReferidos = listaReferidosCompra.concat(listaReferidosAlquiler)
    .sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

  return {
    error: false,
    nombre: embajador.Nombre + ' ' + embajador.Apellido,
    codigoReferido: embajador.Codigo_Referido,
    nivel: embajador.Nivel || 'Bronce',
    score: score,
    mensajePersonalizado: embajador.Mensaje_Personalizado_IA || '',
    cuentaBanco: embajador.Cuenta_Banco || '', cuentaTipo: embajador.Cuenta_Tipo || '',
    cuentaNumero: embajador.Cuenta_Numero || '', cuentaTitular: embajador.Cuenta_Titular || '',
    cuentaBancariaPlataforma: obtenerCuentaPlataformaFormateada_(),
    premium: {
      activo: esPremiumActivo_(embajador),
      fechaVencimiento: embajador.Fecha_Vencimiento_Premium || '',
      pagoReportado: embajador.Pago_Premium_Reportado === 'Sí',
      cuota: Number(getConfig_('CUOTA_PREMIUM_EMBAJADOR_RD', 300)),
      pctPoolPremium: Number(getConfig_('PCT_EMBAJADOR_DEL_POOL_PREMIUM', 90)),
      pctPoolNormal: Number(getConfig_('PCT_EMBAJADOR_DEL_POOL', 80)),
      scoreMinimoBonoPremium: Number(getConfig_('SCORE_MINIMO_BONO_EMBAJADOR_PREMIUM', 50))
    },
    comision: {
      pctBasePool: pctBasePool, pctEmbajadorBase: pctEmbajadorBase,
      pctEmbajadorEfectivo: pctEmbajadorEfectivo, calificaParaBono: calificaParaBono,
      scoreMinimoBono: scoreMinimoBono, pctBonoAltoDesempeno: pctBonoAltoDesempeno
    },
    kpis: { totalReferidos, validados, cerrados, comisionAcumulada, comisionReservadaMin, comisionReservadaMax },
    referidos: listaReferidos,
    pasos: pasos
  };
}

/**
 * Ejecuta una acción de "mantenimiento propio" desde el Portal/Panel — SOLO
 * una lista blanca de campos seguros (teléfono, correo, disponibilidad).
 * Nunca permite tocar campos calculados (Score, comisiones, IDs, etc.) —
 * esos siguen protegidos exactamente igual que en la hoja de cálculo.
 */
function ejecutarAccionMantenimiento_(e, accion) {
  const codigoPanel = (e.parameter.panel || '').trim();
  const codigoPortal = (e.parameter.portal || '').trim();
  const codigoPortalAlquiler = (e.parameter.portalAlquiler || '').trim();
  const codigoPortalPropietario = (e.parameter.portalPropietario || '').trim();
  const codigoPanelAdmin = (e.parameter.panelAdmin || '').trim();
  const valor = (e.parameter.valor || '').trim();

  // --- Acciones del ADMINISTRADOR (Panel remoto) ---
  if (codigoPanelAdmin) return ejecutarAccionAdmin_(e);

  // --- Acciones del AGENTE ---
  if (codigoPanel) {
    const filaAgente = buscarFilaAgentePorCodigoPanel_(codigoPanel);
    if (filaAgente === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo (revisa que no se haya cortado al copiarlo), o pide que te reenvíen tu Panel desde el menú del sistema.' };
    const sheet = sh_(SHEETS.AGENTES);

    if (accion === 'toggleDisponibilidad') {
      if (valor !== 'Sí' && valor !== 'No') return { error: true, mensaje: 'Valor no válido.' };
      sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Disponibilidad')).setValue(valor);
      bitacora_('Agente', 'Disponibilidad actualizada desde el Panel', '', valor);
      // Si se reactivó ("Sí"), programa el reintento de asignación para
      // dentro de unos segundos — NO se ejecuta aquí mismo (eso hacía que
      // el botón se sintiera lento, esperando a que terminara de revisar
      // todos los pendientes antes de responder).
      if (valor === 'Sí') programarReintentoAsignacionAsync_('reintentarAsignacionPendientesTrigger_');
      return { error: false }; // el frontend ya recarga el panel por su cuenta — evita recalcularlo 2 veces
    }
    if (accion === 'toggleAtiendeAlquileres') {
      if (valor !== 'Sí' && valor !== 'No') return { error: true, mensaje: 'Valor no válido.' };
      sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Atiende_Alquileres')).setValue(valor);
      bitacora_('Agente', 'Atiende_Alquileres actualizado desde el Panel', '', valor);
      if (valor === 'Sí') programarReintentoAsignacionAsync_('reintentarAsignacionAlquilerPendientesTrigger_');
      return { error: false }; // el frontend ya recarga el panel por su cuenta — evita recalcularlo 2 veces
    }

    // ============================================================
    //  ACCIONES RÁPIDAS DEL PANEL (v52): registrar visita y cierre
    //  directamente desde el Panel de Agente, sin pasar por un
    //  formulario de Google aparte. El ID del comprador/solicitud llega
    //  desde el propio Panel (el agente lo eligió de su lista) — nunca
    //  se escribe a mano, eliminando la clase de errores de "ID mal
    //  escrito" que causó varios problemas reales hoy.
    // ============================================================
    if (accion === 'registrarVisitaAgente') {
      const idComprador = (e.parameter.idComprador || '').trim().toUpperCase();
      if (!idComprador) return { error: true, mensaje: 'Falta el comprador.' };
      const idAgenteActual = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'ID_Agente')).getValue();
      const resultado = registrarResultadoVisita_(
        idComprador, idAgenteActual,
        (e.parameter.quePaso || '').trim(),
        (e.parameter.motivoRechazo || '').trim(),
        Number(e.parameter.compatibilidad) || 0,
        (e.parameter.proximoPaso || '').trim()
      );
      return resultado; // ya trae { error:false, mensaje } — antes se descartaba el mensaje real y se recalculaba todo el panel sin necesidad
    }
    if (accion === 'registrarCierreVentaAgente') {
      const idComprador = (e.parameter.idComprador || '').trim().toUpperCase();
      if (!idComprador) return { error: true, mensaje: 'Falta el comprador.' };
      const idAgenteActual = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'ID_Agente')).getValue();
      const resultado = procesarNuevoCierre(null, {
        idComprador: idComprador, idAgente: idAgenteActual,
        montoVenta: e.parameter.montoVenta, evidenciaUrl: e.parameter.evidenciaUrl,
        notaAgenteSobreMonto: e.parameter.nota || '', fechaCierre: new Date()
      });
      return resultado || { error: true, mensaje: 'No se pudo registrar el cierre.' };
    }
    if (accion === 'registrarVisitaAlquilerAgente') {
      const idSolicitud = (e.parameter.idSolicitud || '').trim().toUpperCase();
      if (!idSolicitud) return { error: true, mensaje: 'Falta la solicitud de alquiler.' };
      const idAgenteActual = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'ID_Agente')).getValue();
      const resultado = registrarVisitaAlquiler_(idSolicitud, idAgenteActual);
      return resultado;
    }
    if (accion === 'registrarCierreAlquilerAgente') {
      const idSolicitud = (e.parameter.idSolicitud || '').trim().toUpperCase();
      if (!idSolicitud) return { error: true, mensaje: 'Falta la solicitud de alquiler.' };
      const idAgenteActual = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'ID_Agente')).getValue();
      const resultado = procesarNuevoCierreAlquiler(null, {
        idSolicitud: idSolicitud, idAgente: idAgenteActual,
        montoMensual: e.parameter.montoMensual, mesesContrato: e.parameter.mesesContrato,
        evidenciaUrl: e.parameter.evidenciaUrl, notaAgenteSobreMonto: e.parameter.nota || '', fechaCierre: new Date()
      });
      return resultado || { error: true, mensaje: 'No se pudo registrar el cierre.' };
    }
    if (accion === 'reportarPagoRealizado') {
      const idCierre = (e.parameter.idCierre || '').trim().toUpperCase();
      const tipoCierre = (e.parameter.tipoCierre || 'Compra').trim();
      const sheetKey = tipoCierre === 'Alquiler' ? SHEETS.CIERRES_ALQUILER : SHEETS.CIERRES;
      const colIdCierre = tipoCierre === 'Alquiler' ? 'ID_Cierre_Alquiler' : 'ID_Cierre';
      const filaCierre = buscarFilaPorId_(sheetKey, idCierre);
      if (filaCierre === -1) return { error: true, mensaje: 'No encontramos ese cierre.' };

      const sheetCierres = sh_(sheetKey);
      const estadoActual = sheetCierres.getRange(filaCierre, col_(sheetKey, 'Estado_Pago_Comision')).getValue();
      if (estadoActual === 'Pagado') return { error: true, mensaje: 'Esta comisión ya está confirmada como pagada.' };

      sheetCierres.getRange(filaCierre, col_(sheetKey, 'Estado_Pago_Comision')).setValue('Pago reportado por agente');
      bitacora_('Agente', 'Envió solicitud de pago de comisión', idCierre, '');

      const emailAdmin = getConfig_('EMAIL_ADMIN', '');
      if (emailAdmin) {
        enviarCorreo_(emailAdmin, '💰 Solicitud de pago de comisión — ' + idCierre,
          '<p>El agente envió una solicitud de pago para el cierre <b>' + idCierre + '</b>, indicando que ya transfirió la comisión. ' +
          'Verifica que el pago llegó y, si todo está en orden, aprueba cambiando el estado a "Pagado" en la hoja "' + sheetKey + '" — el agente recibirá la confirmación automáticamente.</p>');
      }
      return { error: false, mensaje: 'Tu solicitud de pago fue enviada.' };
    }
    if (accion === 'reportarPagoPremium') {
      const idAgenteActual = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'ID_Agente')).getValue();
      return reportarPagoPremium_('Agente', idAgenteActual); // ya trae { error, mensaje } — el frontend recarga el panel por su cuenta
    }
    if (accion === 'actualizarTelefono') {
      if (!valor) return { error: true, mensaje: 'El teléfono no puede estar vacío.' };
      sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Telefono')).setValue(valor);
      const nombreAgente = sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Nombre')).getValue();
      sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Link_WhatsApp'))
        .setValue(linkWhatsApp_(valor, 'Hola ' + nombreAgente + ', te escribo de ' + APP_NAME + '.'));
      bitacora_('Agente', 'Teléfono actualizado desde el Panel', '', valor);
      return { error: false };
    }
    if (accion === 'actualizarEmail') {
      if (!valor || valor.indexOf('@') === -1) return { error: true, mensaje: 'Correo no válido.' };
      sheet.getRange(filaAgente, col_(SHEETS.AGENTES, 'Email')).setValue(valor);
      bitacora_('Agente', 'Correo actualizado desde el Panel', '', valor);
      return { error: false };
    }
    return { error: true, mensaje: 'Acción no reconocida.' };
  }

  // --- Acciones del COMPRADOR ---
  if (codigoPortal) {
    const filaComprador = buscarFilaPorCodigoPortal_(codigoPortal);
    if (filaComprador === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo (revisa que no se haya cortado al copiarlo), o pide que te reenvíen tu Portal desde el menú del sistema.' };
    const sheet = sh_(SHEETS.COMPRADORES);

    if (accion === 'actualizarTelefono') {
      if (!valor) return { error: true, mensaje: 'El teléfono no puede estar vacío.' };
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Telefono')).setValue(valor);
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Telefono_Normalizado')).setValue(normalizarTelefono_(valor));
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Link_WhatsApp')).setValue(linkWhatsApp_(valor, ''));
      bitacora_('Comprador', 'Teléfono actualizado desde el Portal', '', valor);
      return { error: false };
    }
    if (accion === 'actualizarEmail') {
      if (!valor || valor.indexOf('@') === -1) return { error: true, mensaje: 'Correo no válido.' };
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Email')).setValue(valor);
      bitacora_('Comprador', 'Correo actualizado desde el Portal', '', valor);
      return { error: false };
    }
    if (accion === 'actualizarDocumento') {
      const camposValidos = ['Doc_Cedula', 'Doc_Carta_Trabajo', 'Doc_Estados_Cuenta'];
      const campo = (e.parameter.campo || '').trim();
      if (!camposValidos.includes(campo)) return { error: true, mensaje: 'Documento no reconocido.' };
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, campo)).setValue(valor === 'Sí' ? 'Sí' : 'No');
      return { error: false };
    }
    if (accion === 'calificarVisita') {
      const camposValidos = ['Agente_Puntual', 'Agente_Explico_Bien', 'Recomendaria_Agente', 'Comprador_Le_Gusto', 'Comprador_Confirmo_Visita'];
      const campo = (e.parameter.campo || '').trim();
      if (!camposValidos.includes(campo)) return { error: true, mensaje: 'Pregunta no reconocida.' };
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      const visitasSheet = sh_(SHEETS.VISITAS);
      const lastRowVisitas = visitasSheet.getLastRow();
      if (lastRowVisitas > 1) {
        const dataVisitas = visitasSheet.getRange(2, 1, lastRowVisitas - 1, HEADERS[SHEETS.VISITAS].length).getValues();
        for (let i = dataVisitas.length - 1; i >= 0; i--) {
          if (dataVisitas[i][col_(SHEETS.VISITAS, 'ID_Comprador') - 1] === idComprador) {
            visitasSheet.getRange(i + 2, col_(SHEETS.VISITAS, campo)).setValue(valor === 'Sí' ? 'Sí' : 'No');
            bitacora_('Comprador', 'Calificación de visita registrada', idComprador, campo + ' = ' + valor);
            // El comprador dice que la visita NO ocurrió — no se revierte
            // la exclusividad automáticamente (evita un deshacer complejo
            // y riesgoso), pero sí queda marcado para revisión manual y el
            // agente acumula un registro permanente — control real contra
            // fraude, sin bloquear al 95%+ de agentes honestos.
            if (campo === 'Comprador_Confirmo_Visita' && valor === 'No') {
              const idAgenteVisita = dataVisitas[i][col_(SHEETS.VISITAS, 'ID_Agente') - 1];
              marcarPosibleFraudeVisita_(SHEETS.COMPRADORES, filaComprador, idComprador, idAgenteVisita, 'comprador');
            }
            break;
          }
        }
      }
      return { error: false }; // se llama hasta 4 veces por envío de encuesta — antes recalculaba el Portal completo cada vez
    }
    if (accion === 'responderConexionDirecta') {
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      return responderConexionDirecta_((e.parameter.idConexion || '').trim().toUpperCase(), valor, idComprador);
    }
    if (accion === 'confirmarCierreComprador') {
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      return registrarConfirmacionCierreComprador_(idComprador, valor === 'Sí' ? 'Sí' : 'No');
    }
    if (accion === 'solicitarConexionComprador') {
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      return solicitarConexionDesdeComprador_(idComprador, (e.parameter.idPropiedad || '').trim().toUpperCase(), false);
    }
    if (accion === 'registrarVisitaConexionDirecta') {
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      return registrarVisitaConexionDirecta_((e.parameter.idConexion || '').trim().toUpperCase(), idComprador);
    }
    if (accion === 'solicitarReasignacion') {
      const idComprador = sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'ID_Comprador')).getValue();
      const compradorObj = filaAObjeto_(SHEETS.COMPRADORES, sheet.getRange(filaComprador, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
      const idAgenteActual = compradorObj.Agente_Asignado_ID;
      if (!idAgenteActual) return { error: true, mensaje: 'No tienes un agente asignado actualmente.' };

      // Regla simple y determinística (no IA): excluye al agente actual y
      // usa el mismo motor de matching de siempre para elegir el próximo.
      const nuevoAgente = seleccionarMejorAgente_(compradorObj, idAgenteActual);
      if (!nuevoAgente) {
        bitacora_('Comprador', 'Solicitó reasignación — sin otro agente disponible', idComprador, '');
        return { error: false, mensaje: 'Por ahora no hay otro agente disponible en tu zona. Nuestro equipo te contactará pronto.' };
      }

      // Libera el cupo del agente anterior
      const filaAgenteActual = buscarFilaPorId_(SHEETS.AGENTES, idAgenteActual);
      if (filaAgenteActual > -1) {
        const agentesSheet = sh_(SHEETS.AGENTES);
        const actualObj = filaAObjeto_(SHEETS.AGENTES, agentesSheet.getRange(filaAgenteActual, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);
        agentesSheet.getRange(filaAgenteActual, col_(SHEETS.AGENTES, 'Compradores_Asignados_Activos')).setValue(Math.max(0, Number(actualObj.Compradores_Asignados_Activos || 0) - 1));
        if (actualObj.Email) {
          enviarCorreo_(actualObj.Email, 'Comprador reasignado — ' + idComprador,
            '<p>El comprador <b>' + idComprador + '</b> solicitó cambio de agente y fue reasignado. Ya no está bajo tu cartera.</p>');
        }
      }

      // Asigna al nuevo agente (mismo mecanismo que asignarAgenteAComprador_)
      asignarAgenteAComprador_(idComprador, compradorObj, { agentePreseleccionado: nuevoAgente });

      bitacora_('Comprador', 'Reasignado a otro agente por solicitud propia', idComprador, 'De ' + idAgenteActual + ' a ' + nuevoAgente.obj.ID_Agente);
      return { error: false, mensaje: '¡Listo! Te asignamos con ' + nuevoAgente.obj.Nombre + ' ' + nuevoAgente.obj.Apellido + '. Te contactará pronto.' };
    }
    if (accion === 'solicitarAyuda') {
      const tiposValidos = ['Tengo dudas', 'Quiero cambiar de zona', 'Quiero cambiar de agente', 'Quiero reagendar'];
      if (!tiposValidos.includes(valor)) return { error: true, mensaje: 'Tipo de ayuda no reconocido.' };
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Ultima_Solicitud_Ayuda_Tipo')).setValue(valor);
      sheet.getRange(filaComprador, col_(SHEETS.COMPRADORES, 'Ultima_Solicitud_Ayuda_Fecha')).setValue(new Date());
      const compradorObj = filaAObjeto_(SHEETS.COMPRADORES, sheet.getRange(filaComprador, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0]);
      // Notifica a admin Y al agente asignado — esto es justo lo que evita
      // llamadas innecesarias: el agente ya sabe qué necesita el comprador
      // antes de tener que preguntarle.
      const destinatarios = [getConfig_('EMAIL_ADMIN', '')];
      if (compradorObj.Agente_Asignado_ID) {
        const filaAgente = buscarFilaPorId_(SHEETS.AGENTES, compradorObj.Agente_Asignado_ID);
        if (filaAgente > -1) {
          const agenteEmail = sh_(SHEETS.AGENTES).getRange(filaAgente, col_(SHEETS.AGENTES, 'Email')).getValue();
          if (agenteEmail) destinatarios.push(agenteEmail);
        }
      }
      const linkPanelAdmin = obtenerLinkPanelAdminSiExiste_();
      destinatarios.filter(Boolean).forEach(email => {
        const esAdmin = email === getConfig_('EMAIL_ADMIN', '');
        enviarCorreo_(email, '🆘 ' + compradorObj.ID_Comprador + ' solicitó ayuda: ' + valor,
          '<p><b>' + compradorObj.Nombre + ' ' + compradorObj.Apellido + '</b> (' + compradorObj.ID_Comprador + ') indicó: <b>' + valor + '</b></p>' +
          (esAdmin && linkPanelAdmin ? '<p><a href="' + linkPanelAdmin + '">Ver en mi Panel de Administrador</a></p>' : ''));
      });
      bitacora_('Comprador', 'Solicitud de ayuda', compradorObj.ID_Comprador, valor);
      return { error: false };
    }
    if (accion === 'confirmarMontoCierre') {
      const idCierre = (e.parameter.idCierre || '').trim().toUpperCase();
      return procesarConfirmacionMontoCierre_(compradorObj, idCierre, valor);
    }
    return { error: true, mensaje: 'Acción no reconocida.' };
  }

  // --- Acciones del INQUILINO (Portal de Alquiler) ---
  if (codigoPortalAlquiler) {
    const filaSolicitud = buscarFilaAlquilerPorCodigoPortal_(codigoPortalAlquiler);
    if (filaSolicitud === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo, o pide que te reenvíen tu Portal desde el menú del sistema.' };

    if (accion === 'calificarVisitaAlquiler') {
      const camposValidos = ['Agente_Puntual', 'Agente_Explico_Bien', 'Recomendaria_Agente', 'Inquilino_Le_Gusto', 'Inquilino_Confirmo_Visita'];
      const campo = (e.parameter.campo || '').trim();
      if (!camposValidos.includes(campo)) return { error: true, mensaje: 'Pregunta no reconocida.' };
      const idSolicitud = sh_(SHEETS.SOLICITUDES_ALQUILER).getRange(filaSolicitud, col_(SHEETS.SOLICITUDES_ALQUILER, 'ID_Solicitud')).getValue();
      const visitasSheet = sh_(SHEETS.VISITAS_ALQUILER);
      const lastRowVisitas = visitasSheet.getLastRow();
      if (lastRowVisitas > 1) {
        const dataVisitas = visitasSheet.getRange(2, 1, lastRowVisitas - 1, HEADERS[SHEETS.VISITAS_ALQUILER].length).getValues();
        for (let i = dataVisitas.length - 1; i >= 0; i--) {
          if (dataVisitas[i][col_(SHEETS.VISITAS_ALQUILER, 'ID_Solicitud') - 1] === idSolicitud) {
            visitasSheet.getRange(i + 2, col_(SHEETS.VISITAS_ALQUILER, campo)).setValue(valor === 'Sí' ? 'Sí' : 'No');
            bitacora_('Inquilino', 'Calificación de visita registrada', idSolicitud, campo + ' = ' + valor);
            if (campo === 'Inquilino_Confirmo_Visita' && valor === 'No') {
              const idAgenteVisita = dataVisitas[i][col_(SHEETS.VISITAS_ALQUILER, 'ID_Agente') - 1];
              marcarPosibleFraudeVisita_(SHEETS.SOLICITUDES_ALQUILER, filaSolicitud, idSolicitud, idAgenteVisita, 'inquilino');
            }
            break;
          }
        }
      }
      return { error: false }; // se llama hasta 4 veces por envío de encuesta — antes recalculaba el Portal completo cada vez
    }
    if (accion === 'confirmarMontoCierre') {
      const idCierre = (e.parameter.idCierre || '').trim().toUpperCase();
      const solicitudObj = filaAObjeto_(SHEETS.SOLICITUDES_ALQUILER, sh_(SHEETS.SOLICITUDES_ALQUILER).getRange(filaSolicitud, 1, 1, HEADERS[SHEETS.SOLICITUDES_ALQUILER].length).getValues()[0]);
      return procesarConfirmacionMontoCierreAlquiler_(solicitudObj, idCierre, valor);
    }
    if (accion === 'responderConexionDirecta') {
      const idSolicitud = sh_(SHEETS.SOLICITUDES_ALQUILER).getRange(filaSolicitud, col_(SHEETS.SOLICITUDES_ALQUILER, 'ID_Solicitud')).getValue();
      return responderConexionDirecta_((e.parameter.idConexion || '').trim().toUpperCase(), valor, idSolicitud);
    }
    return { error: true, mensaje: 'Acción no reconocida.' };
  }

  // --- Acciones del PROPIETARIO (Portal de Propietario) ---
  if (codigoPortalPropietario) {
    const filaPropiedad = buscarFilaPropiedadPorCodigoPortal_(codigoPortalPropietario);
    if (filaPropiedad === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo, o pide que te reenvíen tu Portal desde el menú del sistema.' };
    const idPropiedad = sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'ID_Propiedad')).getValue();

    if (accion === 'solicitarConexionDirecta') {
      return solicitarConexionDirecta_(idPropiedad);
    }
    if (accion === 'reportarVentaConexionDirecta') {
      return reportarVentaConexionDirecta_(idPropiedad, (e.parameter.idConexion || '').trim().toUpperCase(), e.parameter.montoVenta);
    }
    if (accion === 'reportarPagoConexionDirecta') {
      return reportarPagoConexionDirecta_((e.parameter.idConexion || '').trim().toUpperCase());
    }
    if (accion === 'responderConexionDirectaPropietario') {
      return responderConexionDirectaPropietario_((e.parameter.idConexion || '').trim().toUpperCase(), valor, idPropiedad);
    }
    return { error: true, mensaje: 'Acción no reconocida.' };
  }

  // --- Acciones del EMBAJADOR ---
  const codigoPanelEmbajador = (e.parameter.embajador || '').trim();
  if (codigoPanelEmbajador) {
    const filaEmbajador = buscarFilaEmbajadorPorCodigoPanel_(codigoPanelEmbajador);
    if (filaEmbajador === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo, o pide que te reenvíen tu Panel desde el menú del sistema.' };
    const sheet = sh_(SHEETS.EMBAJADORES);

    if (accion === 'regenerarMensaje') {
      const embajadorObj = filaAObjeto_(SHEETS.EMBAJADORES, sheet.getRange(filaEmbajador, 1, 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()[0]);
      const link = construirLinkEmbajador_(embajadorObj.Codigo_Referido);
      const nuevoMensaje = generarMensajeReferidoIA_(embajadorObj.Nombre, embajadorObj.Canal_Principal || 'WhatsApp (contactos directos)', link);
      sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Mensaje_Personalizado_IA')).setValue(nuevoMensaje);
      bitacora_('Embajador', 'Mensaje regenerado desde el Panel', embajadorObj.Codigo_Referido, '');
      return obtenerDatosPanelEmbajadorJSON_(codigoPanelEmbajador);
    }
    if (accion === 'actualizarCuentaBancaria') {
      const banco = (e.parameter.banco || '').trim();
      const numero = (e.parameter.numero || '').trim();
      if (!banco || !numero) return { error: true, mensaje: 'El banco y el número de cuenta son obligatorios.' };
      sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Cuenta_Banco')).setValue(banco);
      sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Cuenta_Tipo')).setValue((e.parameter.tipo || '').trim());
      sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Cuenta_Numero')).setValue(numero);
      sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Cuenta_Titular')).setValue((e.parameter.titular || '').trim());
      bitacora_('Embajador', 'Cuenta bancaria actualizada desde el Panel', '', '');
      return { error: false };
    }
    if (accion === 'reportarPagoPremium') {
      const codigoEmbajadorActual = sheet.getRange(filaEmbajador, col_(SHEETS.EMBAJADORES, 'Codigo_Referido')).getValue();
      const resultado = reportarPagoPremium_('Embajador', codigoEmbajadorActual);
      return resultado.error ? resultado : obtenerDatosPanelEmbajadorJSON_(codigoPanelEmbajador);
    }
    return { error: true, mensaje: 'Acción no reconocida.' };
  }

  return { error: true, mensaje: 'Falta el código de acceso (portal o panel).' };
}

/** Reenvía el link del Panel de Embajador por si lo perdió (menú) */
function reenviarLinkPanelEmbajador() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Reenviar Panel de Embajador', 'Código de Embajador (ej. AB4589):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const codigo = resp.getResponseText().trim().toUpperCase();

  const fila = buscarFilaEmbajadorPorCodigo_(codigo);
  if (fila === -1) { alertaSegura_('', '⚠️ No se encontró el embajador con código ' + codigo + '.'); return; }

  const sheet = sh_(SHEETS.EMBAJADORES);
  let obj = filaAObjeto_(SHEETS.EMBAJADORES, sheet.getRange(fila, 1, 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()[0]);
  if (!obj.Codigo_Panel_Privado) {
    obj.Codigo_Panel_Privado = generarCodigoPrivadoLargo_();
    sheet.getRange(fila, col_(SHEETS.EMBAJADORES, 'Codigo_Panel_Privado')).setValue(obj.Codigo_Panel_Privado);
  }

  const link = construirLinkPanelEmbajador_(obj.Codigo_Panel_Privado);
  if (!link) { alertaSegura_('', '⚠️ El Portal de Resultados no está desplegado todavía.'); return; }

  if (obj.Email) {
    enviarCorreo_(obj.Email, 'Tu Panel de Embajador — ' + codigo,
      '<div style="font-family:Arial"><h2>Aquí está tu Panel de Embajador</h2>' +
      '<p><a href="' + link + '">Ver mi panel</a> — guárdalo, es personal e intransferible.</p></div>');
  }
  mostrarPopupConWhatsApp_('📊 Panel de Embajador reenviado', link, obj.Telefono,
    'Hola ' + obj.Nombre + ', aquí está el link a tu Panel de Embajador: ' + link);
}
