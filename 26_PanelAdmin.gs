/**
 * ============================================================================
 *  PANEL DE ADMINISTRADOR — acceso remoto (celular/cualquier navegador) a
 *  las acciones más usadas del menú, con un modelo de seguridad DISTINTO
 *  y MÁS ESTRICTO que el resto del sistema:
 *
 *  - El código secreto NUNCA se guarda en una celda de hoja (que
 *    cualquier editor del Sheet podría ver) — vive en PropertiesService,
 *    accesible solo desde dentro del editor de Apps Script.
 *  - El código NUNCA aparece en el código público de GitHub — admin.html
 *    no contiene ningún secreto; el secreto solo existe en la URL que tú
 *    guardas, generada una sola vez desde el menú.
 *  - Si el código no coincide EXACTAMENTE, no se revela ninguna pista de
 *    qué falló — solo "acceso denegado".
 * ============================================================================
 */

/** Genera (o muestra, si ya existe) el link secreto y privado del Panel de Administrador — SOLO tú debes verlo */
/**
 * Devuelve el link del Panel de Admin SOLO si el código secreto ya fue
 * generado antes (vía el menú) — nunca lo crea automáticamente desde una
 * notificación de fondo, para no generar códigos secretos sin que el
 * admin lo haya pedido explícitamente él mismo.
 */
function obtenerLinkPanelAdminSiExiste_() {
  const codigo = PropertiesService.getScriptProperties().getProperty('ADMIN_PANEL_SECRET_CODE');
  if (!codigo) return '';
  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  return urlGithub
    ? urlGithub.replace(/\/$/, '') + '/admin.html?codigo=' + encodeURIComponent(codigo)
    : (urlBase ? urlBase + (urlBase.includes('?') ? '&' : '?') + 'panelAdmin=' + encodeURIComponent(codigo) : '');
}

function generarLinkPanelAdmin() {
  const props = PropertiesService.getScriptProperties();
  let codigo = props.getProperty('ADMIN_PANEL_SECRET_CODE');
  if (!codigo) {
    codigo = generarCodigoPrivadoLargo_() + generarCodigoPrivadoLargo_(); // 40 caracteres — el doble de largo que los demás códigos, por ser el más sensible
    props.setProperty('ADMIN_PANEL_SECRET_CODE', codigo);
    bitacora_('Sistema', 'Código secreto del Panel de Administrador generado', '', 'Primera vez');
  }

  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  const link = urlGithub
    ? urlGithub.replace(/\/$/, '') + '/admin.html?codigo=' + encodeURIComponent(codigo)
    : (urlBase ? urlBase + (urlBase.includes('?') ? '&' : '?') + 'panelAdmin=' + encodeURIComponent(codigo) : '');

  alertaSegura_('🔐 Tu Panel de Administrador (privado — NO compartir)',
    (link || 'Configura URL_GITHUB_PAGES o URL_PORTAL_RESULTADOS en Configuración primero.') +
    '\n\nGuárdalo tú mismo (favoritos del navegador, notas privadas) — quien tenga este link tiene acceso completo de administrador.');
}

/** Invalida el código actual y genera uno nuevo — usar si sospechas que el link se filtró */
function regenerarLinkPanelAdmin() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('⚠️ Regenerar código secreto',
    'Esto invalida el link actual del Panel de Administrador — cualquier link viejo (guardado en otro dispositivo) dejará de funcionar. ¿Continuar?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  PropertiesService.getScriptProperties().deleteProperty('ADMIN_PANEL_SECRET_CODE');
  bitacora_('Sistema', 'Código secreto del Panel de Administrador regenerado', '', '');
  generarLinkPanelAdmin();
}

/** Verifica el código contra el ÚNICO secreto guardado — sin dar ninguna pista si falla */
function verificarCodigoAdmin_(codigo) {
  const secreto = PropertiesService.getScriptProperties().getProperty('ADMIN_PANEL_SECRET_CODE');
  return !!secreto && !!codigo && codigo === secreto;
}

/** Datos del Panel de Administrador — resumen ejecutivo + estado de pendientes */
function obtenerDatosPanelAdminJSON_(codigo) {
  if (!verificarCodigoAdmin_(codigo)) return { error: true, mensaje: 'Acceso denegado.' };

  const contarPendientesAcuerdos_ = () => {
    const sheet = sh_(SHEETS.AGENTES);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;
    return sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.AGENTES].length).getValues()
      .map(f => filaAObjeto_(SHEETS.AGENTES, f))
      .filter(o => o.Acuerdo_Comision_URL && o.Telefono && o.Acuerdo_Enviado_WhatsApp !== 'Sí').length;
  };
  const contarPendientesPaneles_ = () => {
    let total = 0;
    [[SHEETS.COMPRADORES, 'Codigo_Portal_Privado'], [SHEETS.AGENTES, 'Codigo_Panel_Privado'], [SHEETS.EMBAJADORES, 'Codigo_Panel_Privado']]
      .forEach(([key, campo]) => {
        const sheet = sh_(key);
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        total += sheet.getRange(2, 1, lastRow - 1, HEADERS[key].length).getValues()
          .map(f => filaAObjeto_(key, f))
          .filter(o => o[campo] && o.Telefono && o.Panel_Enviado_WhatsApp !== 'Sí').length;
      });
    return total;
  };
  const contarComisionesPendientesPago_ = () => {
    let total = 0;
    [SHEETS.CIERRES, SHEETS.CIERRES_ALQUILER].forEach(key => {
      const sheet = sh_(key);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      total += sheet.getRange(2, 1, lastRow - 1, HEADERS[key].length).getValues()
        .map(f => filaAObjeto_(key, f))
        .filter(o => o.Estado_Pago_Comision && o.Estado_Pago_Comision !== 'Pagado').length;
    });
    return total;
  };

  const contarComisionesProximasVencer_ = () => {
    const diasAntes = Number(getConfig_('DIAS_RECORDATORIO_PREVIO_COMISION', 3));
    const hoy = new Date();
    let proximasAVencer = 0, yaVencidas = 0;
    [SHEETS.CIERRES, SHEETS.CIERRES_ALQUILER].forEach(key => {
      const sheet = sh_(key);
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      sheet.getRange(2, 1, lastRow - 1, HEADERS[key].length).getValues().forEach(f => {
        const obj = filaAObjeto_(key, f);
        if (obj.Estado_Pago_Comision !== 'Pendiente' || !obj.Fecha_Limite_Pago) return;
        const diasRestantes = Math.ceil((new Date(obj.Fecha_Limite_Pago) - hoy) / (1000 * 60 * 60 * 24));
        if (diasRestantes < 0) yaVencidas++;
        else if (diasRestantes <= diasAntes) proximasAVencer++;
      });
    });
    return { proximasAVencer, yaVencidas };
  };
  const vencimientos = contarComisionesProximasVencer_();

  const contarEmbajadores_ = () => {
    const sheet = sh_(SHEETS.EMBAJADORES);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { total: 0, activos: 0, comisionAcumuladaTotal: 0 };
    const objetos = sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.EMBAJADORES].length).getValues()
      .map(f => filaAObjeto_(SHEETS.EMBAJADORES, f));
    return {
      total: objetos.length,
      activos: objetos.filter(o => o.Estado === 'Activo').length,
      comisionAcumuladaTotal: objetos.reduce((s, o) => s + (Number(o.Comision_Acumulada_RD) || 0), 0)
    };
  };
  const embajadores = contarEmbajadores_();

  // NUEVO — de las mejoras de hoy: % de satisfacción real, combinando la
  // opinión de compradores e inquilinos sobre la propiedad visitada
  // (Comprador_Le_Gusto / Inquilino_Le_Gusto) — un termómetro directo de
  // qué tan bien está funcionando el matching, no solo cuántas visitas hay.
  const contarSatisfaccion_ = () => {
    let siLeGusto = 0, totalConOpinion = 0;
    [[SHEETS.VISITAS, 'Comprador_Le_Gusto'], [SHEETS.VISITAS_ALQUILER, 'Inquilino_Le_Gusto']].forEach(([key, campo]) => {
      let sheet;
      try { sheet = sh_(key); } catch (e) { return; } // hoja aún no creada (falta correr Parte 1A) — se ignora sin romper el Panel
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return;
      sheet.getRange(2, 1, lastRow - 1, HEADERS[key].length).getValues().forEach(f => {
        const obj = filaAObjeto_(key, f);
        if (obj[campo] === 'Sí') { siLeGusto++; totalConOpinion++; }
        else if (obj[campo] === 'No') { totalConOpinion++; }
      });
    });
    return totalConOpinion ? Math.round((siLeGusto / totalConOpinion) * 100) : null;
  };

  // NUEVO — de las mejoras de hoy: zona con más compradores activos, útil
  // para decidir dónde reforzar agentes o usar la redistribución por zona.
  const obtenerZonaMasActiva_ = () => {
    const sheet = sh_(SHEETS.COMPRADORES);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const conteos = {};
    sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.COMPRADORES].length).getValues().forEach(f => {
      const obj = filaAObjeto_(SHEETS.COMPRADORES, f);
      if (obj.Estado === 'Cerrado' || obj.Estado === 'Perdido' || !obj.Provincia) return;
      conteos[obj.Provincia] = (conteos[obj.Provincia] || 0) + 1;
    });
    const entradas = Object.entries(conteos).sort((a, b) => b[1] - a[1]);
    return entradas.length ? { zona: entradas[0][0], cantidad: entradas[0][1] } : null;
  };
  const zonaMasActiva = obtenerZonaMasActiva_();

  // NUEVO — control de exclusividad: agentes con al menos 1 visita negada
  // por un comprador/inquilino (posible fraude) — un caso aislado puede
  // ser un malentendido, varios casos son un patrón real a investigar.
  const contarAlertasFraude_ = () => {
    const sheet = sh_(SHEETS.AGENTES);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { agentesConAlertas: 0, totalNegadas: 0 };
    let agentesConAlertas = 0, totalNegadas = 0;
    sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.AGENTES].length).getValues().forEach(f => {
      const n = Number(filaAObjeto_(SHEETS.AGENTES, f).Visitas_Negadas_Por_Cliente) || 0;
      if (n > 0) { agentesConAlertas++; totalNegadas += n; }
    });
    return { agentesConAlertas, totalNegadas };
  };
  const alertasFraude = contarAlertasFraude_();

  return {
    error: false,
    kpis: {
      compradoresHoy: sh_(SHEETS.DASHBOARD).getRange('B4').getValue() || 0,
      compradoresActivos: sh_(SHEETS.DASHBOARD).getRange('B11').getValue() || 0,
      cierresTotales: sh_(SHEETS.DASHBOARD).getRange('B12').getValue() || 0,
      agentesActivos: sh_(SHEETS.DASHBOARD).getRange('B14').getValue() || 0,
      satisfaccionPct: contarSatisfaccion_(),
      zonaMasActiva: zonaMasActiva ? (zonaMasActiva.zona + ' (' + zonaMasActiva.cantidad + ')') : '—',
      agentesConAlertasFraude: alertasFraude.agentesConAlertas,
      totalVisitasNegadas: alertasFraude.totalNegadas
    },
    embajadores: {
      total: embajadores.total,
      activos: embajadores.activos,
      comisionAcumuladaTotal: embajadores.comisionAcumuladaTotal
    },
    alquiler: {
      solicitudesHoy: sh_(SHEETS.DASHBOARD).getRange('B21').getValue() || 0,
      activas: sh_(SHEETS.DASHBOARD).getRange('B27').getValue() || 0,
      cierres: sh_(SHEETS.DASHBOARD).getRange('B28').getValue() || 0,
      comisionGenerada: sh_(SHEETS.DASHBOARD).getRange('B32').getValue() || 0
    },
    pendientes: {
      acuerdosComision: contarPendientesAcuerdos_(),
      linksPanel: contarPendientesPaneles_(),
      comisionesPorPagar: contarComisionesPendientesPago_(),
      comisionesProximasVencer: vencimientos.proximasAVencer,
      comisionesYaVencidas: vencimientos.yaVencidas
    },
    confirmacionesCierrePendientes: obtenerConfirmacionesCierrePendientes_(),
    conexionesDirectasPendientes: obtenerConexionesDirectasPendientesParaAdmin_()
  };
}

/**
 * Lista de compradores a quienes se les preguntó si cerraron su compra
 * y todavía no han respondido — con un enlace de WhatsApp listo para
 * que el admin les dé un empujón manual si lo considera necesario.
 */
function obtenerConfirmacionesCierrePendientes_() {
  const sheet = sh_(SHEETS.COMPRADORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
    .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
    .filter(c => c.Cierre_Confirmado_Comprador === 'Pendiente')
    .map(c => ({
      id: c.ID_Comprador,
      nombre: c.Nombre + ' ' + c.Apellido,
      fechaPregunta: c.Fecha_Pregunta_Cierre ? Utilities.formatDate(new Date(c.Fecha_Pregunta_Cierre), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      linkWhatsApp: linkWhatsApp_(c.Telefono, 'Hola ' + c.Nombre + ', te escribimos de ' + APP_NAME + ' — ¿llegaste a cerrar la compra de tu propiedad? Nos ayudaría mucho saberlo 🙏')
    }));
}

/**
 * Lista de solicitudes de Conexión Directa (comprador ↔ propietario)
 * esperando respuesta de alguna de las 2 partes — con un enlace de
 * WhatsApp listo para que el admin le dé un empujón manual a quien le
 * toque responder.
 */
function obtenerConexionesDirectasPendientesParaAdmin_() {
  let sheet;
  try { sheet = sh_(SHEETS.CONEXIONES_DIRECTAS); } catch (e) { return []; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const pendientes = sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()
    .map(f => filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f))
    .filter(c => c.Estado === 'Pendiente_Respuesta_Comprador' || c.Estado === 'Pendiente_Respuesta_Propietario');

  return pendientes.map(c => {
    const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, c.ID_Propiedad);
    const propiedad = filaPropiedad > -1
      ? filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0])
      : null;
    const filaComprador = buscarFilaPorId_(SHEETS.COMPRADORES, c.ID_Comprador);
    const comprador = filaComprador > -1
      ? filaAObjeto_(SHEETS.COMPRADORES, sh_(SHEETS.COMPRADORES).getRange(filaComprador, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0])
      : null;

    const esperandoAlPropietario = c.Estado === 'Pendiente_Respuesta_Propietario';
    const quienDebeResponder = esperandoAlPropietario
      ? (propiedad ? propiedad.Nombre_Propietario : 'Propietario') + ' (propietario)'
      : (comprador ? comprador.Nombre : 'Comprador') + ' (comprador)';
    const telefonoAContactar = esperandoAlPropietario ? (propiedad ? propiedad.Telefono : '') : (comprador ? comprador.Telefono : '');
    const nombreAContactar = esperandoAlPropietario ? (propiedad ? propiedad.Nombre_Propietario : '') : (comprador ? comprador.Nombre : '');

    return {
      idConexion: c.ID_Conexion,
      idPropiedad: c.ID_Propiedad,
      iniciadaPor: c.Iniciada_Por || '—',
      esperandoA: quienDebeResponder,
      fecha: c.Fecha_Solicitud ? Utilities.formatDate(new Date(c.Fecha_Solicitud), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      linkWhatsApp: telefonoAContactar
        ? linkWhatsApp_(telefonoAContactar, 'Hola ' + nombreAContactar + ', te escribimos de ' + APP_NAME + ' — tienes una solicitud de Conexión Directa esperando tu respuesta en tu Portal. ¿Puedes revisarla? 🙏')
        : ''
    };
  });
}

/** Ejecuta una acción administrativa remota — todas requieren el código secreto verificado */
function ejecutarAccionAdmin_(e) {
  const codigo = (e.parameter.panelAdmin || '').trim();
  if (!verificarCodigoAdmin_(codigo)) return { error: true, mensaje: 'Acceso denegado.' };

  const accion = (e.parameter.accion || '').trim();
  bitacora_('Admin', 'Acción ejecutada desde el Panel de Administrador remoto', accion, '');

  if (accion === 'enviarAcuerdosPendientes') {
    const resultado = enviarAcuerdosComisionPendientesPorWhatsApp();
    return { error: false, resultado: resultado };
  }
  if (accion === 'enviarPanelesPendientes') {
    const resultado = enviarPanelesPendientesPorWhatsApp();
    return { error: false, resultado: resultado };
  }
  if (accion === 'verificarIntegridad') {
    const problemas = verificarIntegridadDatos();
    return { error: false, problemas: problemas };
  }
  if (accion === 'actualizarDashboard') {
    actualizarDashboard();
    return { error: false, mensaje: 'Dashboard actualizado.' };
  }
  if (accion === 'marcarEnviado') {
    // Se llama apenas el admin abre el link de WhatsApp desde el Panel —
    // no hay forma de saber con certeza si de verdad presionó "enviar"
    // dentro de WhatsApp (eso ya es fuera del alcance de cualquier
    // sistema), pero abrir el link es la señal más confiable disponible.
    const tipo = (e.parameter.tipo || '').trim();
    const sheetKey = (e.parameter.sheetKey || '').trim();
    const id = (e.parameter.id || '').trim();
    if (!sheetKey || !id) return { error: true, mensaje: 'Faltan datos.' };

    const campoEstado = tipo === 'acuerdo' ? 'Acuerdo_Enviado_WhatsApp' : 'Panel_Enviado_WhatsApp';
    const fila = buscarFilaPorId_(sheetKey, id);
    if (fila === -1) return { error: true, mensaje: 'No se encontró el registro.' };

    sh_(sheetKey).getRange(fila, col_(sheetKey, campoEstado)).setValue('Sí');
    bitacora_('Admin', 'Marcado como enviado por WhatsApp desde el Panel de Administrador', id, tipo);
    return { error: false };
  }
  if (accion === 'redistribuirZona') {
    const zona = (e.parameter.zona || '').trim();
    if (!zona) return { error: true, mensaje: 'Falta la zona.' };
    return calcularRedistribucionPorZona_(zona, true);
  }
  if (accion === 'verComentariosVisitas') return { error: false, comentarios: obtenerComentariosVisitasRecientes_() };
  if (accion === 'verAgentesPendientesVerificar') return { error: false, agentes: obtenerAgentesPendientesVerificar_() };
  if (accion === 'verSolicitudesAyuda') return { error: false, solicitudes: obtenerSolicitudesAyudaRecientes_() };
  if (accion === 'verAlertasFraude') return { error: false, alertas: obtenerDetalleAlertasFraude_() };
  return { error: true, mensaje: 'Acción no reconocida.' };
}

/**
 * Estas 4 funciones son "bajo demanda" — no se cargan cuando abres el
 * Panel de Admin (eso lo haría más lento en cada visita), solo cuando
 * presionas el botón correspondiente. Buscan datos que hoy solo se podían
 * ver abriendo el Sheet directamente.
 */

/** Últimos comentarios de visitas — comprador/inquilino, agente, opinión y fecha, combinando compra y alquiler */
function obtenerComentariosVisitasRecientes_() {
  const resultado = [];

  const procesarHoja_ = (sheetKey, campoId, campoLeGusto, etiquetaTipo) => {
    let sheet;
    try { sheet = sh_(sheetKey); } catch (e) { return; } // hoja aún no creada — se ignora sin romper el Panel
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    sheet.getRange(2, 1, lastRow - 1, HEADERS[sheetKey].length).getValues().forEach(f => {
      const obj = filaAObjeto_(sheetKey, f);
      if (!obj[campoLeGusto]) return; // sin opinión todavía, no mostrar
      const filaAgente = buscarFilaPorId_(SHEETS.AGENTES, obj.ID_Agente);
      const nombreAgente = filaAgente > -1
        ? (() => { const a = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]); return a.Nombre + ' ' + a.Apellido; })()
        : obj.ID_Agente;
      resultado.push({
        tipo: etiquetaTipo, id: obj[campoId], agente: nombreAgente,
        leGusto: obj[campoLeGusto], fecha: obj.Fecha_Registro_Resultado || obj.Fecha_Registro || ''
      });
    });
  };

  procesarHoja_(SHEETS.VISITAS, 'ID_Comprador', 'Comprador_Le_Gusto', 'Comprador');
  procesarHoja_(SHEETS.VISITAS_ALQUILER, 'ID_Solicitud', 'Inquilino_Le_Gusto', 'Inquilino');

  return resultado.sort((a, b) => new Date(b.fecha) - new Date(a.fecha)).slice(0, 20);
}

/** Agentes registrados pero aún sin verificar — pendientes de tu revisión manual */
function obtenerAgentesPendientesVerificar_() {
  const sheet = sh_(SHEETS.AGENTES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.AGENTES].length).getValues()
    .map(f => filaAObjeto_(SHEETS.AGENTES, f))
    .filter(a => a.Verificado !== 'Sí')
    .map(a => ({ id: a.ID_Agente, nombre: a.Nombre + ' ' + a.Apellido, telefono: a.Telefono, fecha: a.Fecha_Registro }));
}

/** Solicitudes de ayuda recientes de compradores — hoy solo llegaban por correo */
function obtenerSolicitudesAyudaRecientes_() {
  const sheet = sh_(SHEETS.COMPRADORES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.COMPRADORES].length).getValues()
    .map(f => filaAObjeto_(SHEETS.COMPRADORES, f))
    .filter(c => c.Ultima_Solicitud_Ayuda_Tipo)
    .map(c => ({ id: c.ID_Comprador, nombre: c.Nombre + ' ' + c.Apellido, tipo: c.Ultima_Solicitud_Ayuda_Tipo, fecha: c.Ultima_Solicitud_Ayuda_Fecha }))
    .sort((a, b) => new Date(b.fecha) - new Date(a.fecha))
    .slice(0, 20);
}

/** Detalle de agentes con visitas negadas por clientes — antes solo se veía el conteo total */
function obtenerDetalleAlertasFraude_() {
  const sheet = sh_(SHEETS.AGENTES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.AGENTES].length).getValues()
    .map(f => filaAObjeto_(SHEETS.AGENTES, f))
    .filter(a => Number(a.Visitas_Negadas_Por_Cliente) > 0)
    .map(a => ({ id: a.ID_Agente, nombre: a.Nombre + ' ' + a.Apellido, cantidad: Number(a.Visitas_Negadas_Por_Cliente) }))
    .sort((a, b) => b.cantidad - a.cantidad);
}
