/**
 * ============================================================================
 *  MÓDULO DE PROPIETARIOS — nueva pieza del marketplace
 * ============================================================================
 *  Conecta propietarios que venden/alquilan directamente (no solo
 *  compradores buscando). Incluye el "Modo Inteligente": el propietario
 *  elige si quiere vender/alquilar solo, con un agente que ya conoce, o
 *  que el sistema le recomiende el mejor agente disponible.
 * ============================================================================
 */

function procesarNuevaPropiedad(e) {
  try {
    if (!e || !e.namedValues) {
      log_('procesarNuevaPropiedad', 'WARN', 'Evento sin namedValues — probablemente vista previa del formulario, no un envío real.');
      return;
    }
    const r = e.namedValues;
    const val = (campo) => (r[campo] && r[campo][0]) ? r[campo][0].trim() : '';

    const telefono = val('Teléfono / WhatsApp');
    const email = val('Correo electrónico');
    const idPropiedad = generarIdSecuencial_('PROP', SHEETS.PROPIEDADES, col_(SHEETS.PROPIEDADES, 'ID_Propiedad'));

    const propiedad = {
      ID_Propiedad: idPropiedad,
      Fecha_Registro: new Date(),
      Nombre_Propietario: val('Nombre'),
      Apellido_Propietario: val('Apellido'),
      Telefono: telefono,
      Telefono_Normalizado: normalizarTelefono_(telefono),
      Email: email,
      Provincia: val('Provincia'),
      Municipio: obtenerMunicipioSeleccionado_(r),
      Tipo_Inmueble: val('Tipo de inmueble'),
      Operacion: val('¿Qué deseas hacer?'),
      Precio: Number(val('Precio (RD$)')) || 0,
      Habitaciones: val('Habitaciones'),
      Banos: val('Baños'),
      Parqueos: val('Parqueos'),
      Descripcion: val('Descripción breve de la propiedad'),
      Modo_Venta: val('¿Cómo deseas vender/alquilar tu propiedad?'),
      Agente_Asignado_ID: '', Agente_Asignado_Nombre: '', Fecha_Asignacion: '',
      Verificado: 'No', Fecha_Verificacion: '',
      Estado: 'Activa',
      Codigo_Portal_Privado: generarCodigoPrivadoLargo_(),
      Codigo_Referido: val('Código de Embajador (no modificar)').toUpperCase(),
      Estado_Referido: '',
      Autorizacion_Contacto: val('Autorizo ser contactado por agentes inmobiliarios verificados').includes('Sí') ? 'Sí' : 'No',
      Consentimiento_Datos: val('Doy mi consentimiento para el tratamiento de mis datos personales (Ley 172-13, RD)').includes('Sí') ? 'Sí' : 'No',
      Fecha_Consentimiento: new Date(),
      Notas: val('¿Cómo deseas vender/alquilar tu propiedad?') === 'Con un agente que ya conozco'
        ? 'Agente indicado por el propietario: ' + val('Código o nombre del agente que ya conoces')
        : ''
    };

    const headers = HEADERS[SHEETS.PROPIEDADES];
    const fila = headers.map(h => propiedad[h] !== undefined ? propiedad[h] : '');
    sh_(SHEETS.PROPIEDADES).appendRow(fila);

    bitacora_('Formulario Propietario', 'Nueva propiedad registrada', idPropiedad,
      propiedad.Operacion + ' — ' + propiedad.Tipo_Inmueble + ' en ' + propiedad.Provincia + ' — Modo: ' + propiedad.Modo_Venta);
    notificarAdminNuevoRegistro_('Propietario', idPropiedad, propiedad.Nombre_Propietario,
      propiedad.Operacion + ' — ' + propiedad.Tipo_Inmueble + ' en ' + propiedad.Provincia + ' · Modo: ' + propiedad.Modo_Venta);

    // Red de Embajadores: ahora también pueden referir propietarios, no
    // solo compradores/inquilinos — el registro de contador es
    // inmediato; la comisión real (cuando aplique) se acredita al marcar
    // la propiedad como Vendida/Alquilada, igual que con un cierre normal.
    if (propiedad.Codigo_Referido) {
      const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, idPropiedad);
      propiedad.Estado_Referido = validarCodigoEmbajador_(propiedad) ? 'Pendiente' : 'Código inválido';
      if (filaPropiedad > -1) {
        sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'Estado_Referido')).setValue(propiedad.Estado_Referido);
      }
      bitacora_('Red de Embajadores', 'Referido de propiedad registrado (' + propiedad.Estado_Referido + ')', idPropiedad, 'Código: ' + propiedad.Codigo_Referido);
      if (propiedad.Estado_Referido === 'Pendiente') {
        actualizarContadorEmbajador_(propiedad.Codigo_Referido, 'Total_Referidos');
      }
    }

    // Distribución según el Modo Inteligente elegido
    if (propiedad.Autorizacion_Contacto === 'Sí' && propiedad.Consentimiento_Datos === 'Sí') {
      if (propiedad.Modo_Venta === 'Que PropMatch me recomiende el mejor agente') {
        asignarMejorAgenteAPropiedad_(idPropiedad, propiedad);
      } else if (propiedad.Modo_Venta === 'Con un agente que ya conozco') {
        // No se asigna automáticamente — requiere que el admin confirme
        // manualmente cuál es el agente exacto (el propietario solo dio un
        // nombre o código posiblemente impreciso), evitando asignar a la
        // persona equivocada por un error de tecleo.
        const emailAdmin = getConfig_('EMAIL_ADMIN', '');
        if (emailAdmin) {
          enviarCorreo_(emailAdmin, '👤 Propietario pidió un agente específico — ' + idPropiedad,
            '<p>El propietario de <b>' + idPropiedad + '</b> indicó que ya tiene un agente: <b>' + propiedad.Notas + '</b>.</p>' +
            '<p>Verifica quién es exactamente y asígnalo manualmente en la hoja "Propiedades" (columna Agente_Asignado_ID).</p>');
        }
      }
      // "Yo solo (sin agente)": la propiedad queda registrada sin agente
      // asignado — el propietario gestiona directamente, la plataforma
      // solo la deja visible para el futuro motor de Instant Match.
    }

    // FASE 2: Instant Match — no se espera a que alguien encuentre el
    // anuncio, se buscan de inmediato compradores/inquilinos compatibles
    // ya registrados en el sistema.
    ejecutarInstantMatch_(idPropiedad, propiedad);

    // Confirmación al propietario
    if (getConfig_('ENVIAR_CORREO_PROPIETARIO', 'Sí') === 'Sí' && email) {
      const linkPortal = construirLinkPortalPropietario_(propiedad.Codigo_Portal_Privado);
      enviarCorreo_(email, '¡Recibimos tu propiedad! — ' + idPropiedad,
        '<div style="font-family:Arial"><h2>¡Gracias, ' + propiedad.Nombre_Propietario + '!</h2>' +
        '<p>Tu propiedad <b>' + idPropiedad + '</b> fue registrada.</p>' +
        (linkPortal ? '<p style="text-align:center"><a href="' + linkPortal + '" style="display:inline-block;background:#1a73e8;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold;font-size:15px">📊 Ver mi Portal</a></p><p style="font-size:12px;color:#999;text-align:center">Guarda este link — es personal e intransferible.</p>' : '') +
        bloqueSoporteWhatsApp_('Hola, tengo una consulta sobre mi propiedad ' + idPropiedad) +
        '</div>');
    }
  } catch (err) {
    log_('procesarNuevaPropiedad', 'ERROR', err.message + ' | ' + err.stack);
  }
}

/**
 * MODO INTELIGENTE: recomienda y asigna el mejor agente para una
 * propiedad — reutiliza el MISMO motor de puntaje que ya asigna agentes a
 * compradores (construirPuntuadorAgente_ en 05_Depuracion_Asignacion.gs),
 * adaptado para puntuar según la propiedad en vez de según un comprador.
 * Prioriza: coincidencia de zona, tipo de inmueble que maneja, e
 * historial de cierres/tasa de conversión (mismo criterio que ya usa el
 * sistema para priorizar buenos agentes, sin inventar un motor paralelo).
 */
function asignarMejorAgenteAPropiedad_(idPropiedad, propiedad) {
  // Se arma un objeto "similar a comprador" para poder reutilizar el
  // puntuador existente sin duplicar su lógica de puntaje.
  const compradorSimulado = { Provincia: propiedad.Provincia, Tipo_Inmueble: propiedad.Tipo_Inmueble };
  const mejorAgente = seleccionarMejorAgente_(compradorSimulado, null);

  if (!mejorAgente) {
    log_('asignarMejorAgenteAPropiedad_', 'WARN', 'Sin agentes elegibles para la propiedad ' + idPropiedad);
    bitacora_('Sistema', 'Sin agente disponible (propiedad)', idPropiedad, 'No hay agentes verificados/activos/disponibles que cumplan criterios');
    return;
  }

  const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, idPropiedad);
  if (filaPropiedad === -1) return;
  const sheet = sh_(SHEETS.PROPIEDADES);
  sheet.getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'Agente_Asignado_ID')).setValue(mejorAgente.obj.ID_Agente);
  sheet.getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'Agente_Asignado_Nombre')).setValue(mejorAgente.obj.Nombre + ' ' + mejorAgente.obj.Apellido);
  sheet.getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'Fecha_Asignacion')).setValue(new Date());

  bitacora_('Modo Inteligente', 'Agente recomendado y asignado', idPropiedad, mejorAgente.obj.ID_Agente);

  if (mejorAgente.obj.Email) {
    enviarCorreo_(mejorAgente.obj.Email, '🎯 Propiedad recomendada para ti — ' + idPropiedad,
      '<div style="font-family:Arial"><h2>🎯 El Modo Inteligente te recomendó para esta propiedad</h2>' +
      '<p>Un propietario registró su propiedad y nuestro sistema te seleccionó como el agente con mayor probabilidad de éxito para ella, según tu zona, especialidad e historial de cierres.</p>' +
      '<p><b>' + propiedad.Operacion + ':</b> ' + propiedad.Tipo_Inmueble + ' en ' + propiedad.Provincia + ' - ' + propiedad.Municipio + '</p>' +
      '<p><b>Precio:</b> RD$' + Number(propiedad.Precio).toLocaleString() + ' · <b>Habitaciones:</b> ' + propiedad.Habitaciones + ' · <b>Baños:</b> ' + propiedad.Banos + '</p>' +
      (propiedad.Descripcion ? '<p><b>Descripción del propietario:</b> ' + propiedad.Descripcion + '</p>' : '') +
      '</div>');
  }
}

/**
 * PROPMATCH VERIFIED OWNER — mismo patrón exacto que la verificación de
 * agentes: el admin marca Verificado=Sí manualmente en la hoja
 * "Propiedades" tras confirmar identidad y título de propiedad, y esta
 * función (llamada desde el menú) registra la fecha. Reduce fraude al
 * dar más confianza a compradores/inquilinos y agentes de que la
 * propiedad es legítima.
 */
function verificarPropietariosMarcados() {
  const sheet = sh_(SHEETS.PROPIEDADES);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { alertaSegura_('', 'No hay propiedades registradas todavía.'); return; }
  const headers = HEADERS[SHEETS.PROPIEDADES];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const colFechaVerif = col_(SHEETS.PROPIEDADES, 'Fecha_Verificacion');
  let actualizados = 0;

  data.forEach((fila, i) => {
    const obj = filaAObjeto_(SHEETS.PROPIEDADES, fila);
    if (obj.Verificado === 'Sí' && !obj.Fecha_Verificacion) {
      sheet.getRange(i + 2, colFechaVerif).setValue(new Date());
      bitacora_('Admin', 'Propiedad verificada (Verified Owner)', obj.ID_Propiedad, '');
      if (obj.Email) {
        enviarCorreo_(obj.Email, '✅ ¡Tu propiedad fue verificada! — ' + obj.ID_Propiedad,
          '<div style="font-family:Arial"><h2>✅ Insignia PropMatch Verified Owner</h2>' +
          '<p>Verificamos tu identidad y documentación — tu propiedad ahora muestra el sello de confianza para compradores/inquilinos y agentes.</p></div>');
      }
      actualizados++;
    }
  });
  alertaSegura_('', '✅ ' + actualizados + ' propiedad(es) verificadas.');
}

function construirLinkPortalPropietario_(codigo) {
  if (!codigo) return '';
  const urlGithub = getConfig_('URL_GITHUB_PAGES', '');
  const urlBase = getConfig_('URL_PORTAL_RESULTADOS', '');
  return urlGithub
    ? urlGithub.replace(/\/$/, '') + '/portal-propietario.html?codigo=' + encodeURIComponent(codigo)
    : (urlBase ? urlBase + (urlBase.includes('?') ? '&' : '?') + 'portalPropietario=' + encodeURIComponent(codigo) : '');
}

/**
 * ============================================================================
 *  FASE 2: PROPMATCH INSTANT MATCH
 * ============================================================================
 *  Cuando entra una propiedad nueva, no se espera a que alguien la
 *  encuentre — se buscan de inmediato compradores/inquilinos YA
 *  registrados que sean compatibles (misma zona, tipo de inmueble,
 *  presupuesto dentro de rango) y se notifica al agente asignado (o al
 *  propietario, si eligió "Yo solo"). Reutiliza los mismos campos de
 *  matching que ya existen — sin motor de IA aparte, solo cruza los
 *  datos que ya se recolectan de ambos lados del mercado.
 * ============================================================================
 */
function buscarCompradoresCompatiblesConPropiedad_(propiedad) {
  const compatibles = [];
  const esVenta = propiedad.Operacion === 'Vender';
  const sheetKey = esVenta ? SHEETS.COMPRADORES : SHEETS.SOLICITUDES_ALQUILER;
  const campoPrecioMin = esVenta ? 'Presupuesto_Min' : 'Presupuesto_Min_Mensual';
  const campoPrecioMax = esVenta ? 'Presupuesto_Max' : 'Presupuesto_Max_Mensual';
  const campoEstadoDescartar = esVenta ? ['Cerrado', 'Perdido'] : ['Cerrado', 'Perdido'];

  let sheet;
  try { sheet = sh_(sheetKey); } catch (e) { return compatibles; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return compatibles;

  sheet.getRange(2, 1, lastRow - 1, HEADERS[sheetKey].length).getValues().forEach(f => {
    const obj = filaAObjeto_(sheetKey, f);
    if (campoEstadoDescartar.includes(obj.Estado)) return;
    // Coincide con la zona principal del comprador/inquilino O con su 2da zona de interés (si tiene)
    const coincideZona = obj.Provincia === propiedad.Provincia || (obj.Provincia_2 && obj.Provincia_2 === propiedad.Provincia);
    if (!coincideZona) return;
    if (obj.Tipo_Inmueble && propiedad.Tipo_Inmueble && obj.Tipo_Inmueble !== propiedad.Tipo_Inmueble) return;

    const min = Number(obj[campoPrecioMin]) || 0;
    const max = Number(obj[campoPrecioMax]) || Infinity;
    // Margen del 10% para no descartar coincidencias casi perfectas por muy poco
    if (propiedad.Precio < min * 0.9 || propiedad.Precio > max * 1.1) return;

    compatibles.push({ id: obj.ID_Comprador || obj.ID_Solicitud, nombre: obj.Nombre + ' ' + obj.Apellido, telefono: obj.Telefono });
  });

  return compatibles;
}

/**
 * Se llama justo después de registrar una propiedad nueva — busca
 * compradores/inquilinos compatibles y avisa a quien corresponda
 * (agente asignado, o al propio propietario si eligió gestionar solo).
 */
function ejecutarInstantMatch_(idPropiedad, propiedad) {
  const compatibles = buscarCompradoresCompatiblesConPropiedad_(propiedad);
  if (!compatibles.length) return;

  bitacora_('Instant Match', compatibles.length + ' comprador(es)/inquilino(s) compatible(s) encontrados', idPropiedad, compatibles.map(c => c.id).join(', '));

  const listaHtml = compatibles.slice(0, 10).map(c => '<li>' + c.nombre + ' (' + c.id + ')</li>').join('');
  const cuerpo = '<div style="font-family:Arial"><h2>⚡ Instant Match — ya hay ' + compatibles.length + ' interesado(s) esperando</h2>' +
    '<p>Encontramos ' + compatibles.length + ' comprador(es)/inquilino(s) ya registrados que podrían encajar con esta propiedad (' + idPropiedad + '), sin esperar a que alguien la encuentre por su cuenta.</p>' +
    '<ul>' + listaHtml + '</ul>' +
    (compatibles.length > 10 ? '<p>… y ' + (compatibles.length - 10) + ' más.</p>' : '') +
    '</div>';

  if (propiedad.Agente_Asignado_ID) {
    const filaAgente = buscarFilaPorId_(SHEETS.AGENTES, propiedad.Agente_Asignado_ID);
    if (filaAgente > -1) {
      const agenteObj = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);
      if (agenteObj.Email) enviarCorreo_(agenteObj.Email, '⚡ Instant Match — ' + idPropiedad, cuerpo);
    }
  } else if (propiedad.Email) {
    enviarCorreo_(propiedad.Email, '⚡ ¡Ya hay interesados en tu propiedad! — ' + idPropiedad, cuerpo);
  }
}

/**
 * ============================================================================
 *  PORTAL DEL PROPIETARIO — mismo patrón que el Portal de Alquiler
 * ============================================================================
 */
function buscarFilaPropiedadPorCodigoPortal_(codigo) {
  let sheet;
  try { sheet = sh_(SHEETS.PROPIEDADES); } catch (e) { return -1; } // hoja aún no creada — no debe romper el Portal
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const codigoBuscado = String(codigo).trim();
  const codigos = sheet.getRange(2, col_(SHEETS.PROPIEDADES, 'Codigo_Portal_Privado'), lastRow - 1, 1).getValues();
  for (let i = 0; i < codigos.length; i++) {
    if (String(codigos[i][0]).trim() === codigoBuscado) return i + 2;
  }
  return -1;
}

/**
 * Versión ANONIMIZADA de buscarCompradoresCompatiblesConPropiedad_, para
 * mostrar en el Portal del PROPIETARIO — nunca expone nombre, teléfono ni
 * ningún dato personal (el propietario no es el agente asignado, no debe
 * tener contacto directo). Solo cuenta cuántos hay y su clasificación,
 * mismo criterio de anonimización que ya se usa en el Panel de Embajador.
 */
function obtenerResumenCompradoresCompatibles_(propiedad) {
  const esVenta = propiedad.Operacion === 'Vender';
  const sheetKey = esVenta ? SHEETS.COMPRADORES : SHEETS.SOLICITUDES_ALQUILER;
  const campoPrecioMin = esVenta ? 'Presupuesto_Min' : 'Presupuesto_Min_Mensual';
  const campoPrecioMax = esVenta ? 'Presupuesto_Max' : 'Presupuesto_Max_Mensual';

  let sheet;
  try { sheet = sh_(sheetKey); } catch (e) { return { total: 0, porClasificacion: {} }; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { total: 0, porClasificacion: {} };

  const porClasificacion = {};
  let total = 0;

  sheet.getRange(2, 1, lastRow - 1, HEADERS[sheetKey].length).getValues().forEach(f => {
    const obj = filaAObjeto_(sheetKey, f);
    if (['Cerrado', 'Perdido'].includes(obj.Estado)) return;
    const coincideZona = obj.Provincia === propiedad.Provincia || (obj.Provincia_2 && obj.Provincia_2 === propiedad.Provincia);
    if (!coincideZona) return;
    if (obj.Tipo_Inmueble && propiedad.Tipo_Inmueble && obj.Tipo_Inmueble !== propiedad.Tipo_Inmueble) return;
    const min = Number(obj[campoPrecioMin]) || 0;
    const max = Number(obj[campoPrecioMax]) || Infinity;
    if (propiedad.Precio < min * 0.9 || propiedad.Precio > max * 1.1) return;

    total++;
    const clasif = obj.Clasificacion || 'Sin clasificar';
    porClasificacion[clasif] = (porClasificacion[clasif] || 0) + 1;
  });

  return { total, porClasificacion };
}

/**
 * Análisis breve de la propiedad para el propietario — posiciona el
 * precio contra la referencia de mercado (si existe para esa
 * provincia+tipo) y menciona cuántos compradores/inquilinos compatibles
 * ya hay esperando. Mismo espíritu que el Resumen Ejecutivo del
 * comprador, pero desde el lado de la oferta.
 */
function generarAnalisisPropiedad_(propiedad, resumenCompatibles) {
  const partes = [];

  // Posicionamiento de precio vs. referencia de mercado (si hay dato)
  let sheetRef;
  try { sheetRef = sh_(SHEETS.REFERENCIA_PRECIOS); } catch (e) { sheetRef = null; }
  if (sheetRef) {
    const lastRow = sheetRef.getLastRow();
    if (lastRow > 1) {
      const referencia = sheetRef.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.REFERENCIA_PRECIOS].length).getValues()
        .map(f => filaAObjeto_(SHEETS.REFERENCIA_PRECIOS, f))
        .find(r => r.Provincia === propiedad.Provincia && r.Tipo_Inmueble === propiedad.Tipo_Inmueble);
      if (referencia) {
        const min = Number(referencia.Precio_Min_RD), max = Number(referencia.Precio_Max_RD);
        if (propiedad.Precio < min) partes.push('Tu precio está por debajo del rango típico de mercado para tu zona y tipo de inmueble (RD$' + min.toLocaleString() + ' - RD$' + max.toLocaleString() + ') — podrías tener margen para ajustarlo.');
        else if (propiedad.Precio > max) partes.push('Tu precio está por encima del rango típico de mercado para tu zona y tipo de inmueble (RD$' + min.toLocaleString() + ' - RD$' + max.toLocaleString() + ') — esto puede alargar el tiempo de venta/alquiler.');
        else partes.push('Tu precio está dentro del rango típico de mercado para tu zona y tipo de inmueble.');
      }
    }
  }

  // Demanda compatible ya existente
  if (resumenCompatibles.total > 0) {
    const clasifTexto = Object.entries(resumenCompatibles.porClasificacion)
      .sort((a, b) => b[1] - a[1])
      .map(([c, n]) => n + ' ' + c).join(', ');
    partes.push('Ya hay ' + resumenCompatibles.total + ' comprador(es)/inquilino(s) compatible(s) esperando en tu zona (' + clasifTexto + ').');
  } else {
    partes.push('Por ahora no hay compradores/inquilinos compatibles registrados en tu zona todavía — te avisaremos apenas aparezca uno (Instant Match).');
  }

  return partes.join(' ');
}

function obtenerDatosPortalPropietarioJSON_(codigoPortal) {
  const fila = buscarFilaPropiedadPorCodigoPortal_(codigoPortal);
  if (fila === -1) return { error: true, mensaje: 'Enlace no válido. Verifica que copiaste el link completo, o pide que te reenvíen tu Portal desde el menú del sistema.' };

  const obj = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(fila, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);

  let estadoExplicado = 'Tu propiedad fue recibida y está activa en el sistema.';
  if (obj.Estado === 'Vendida') estadoExplicado = '¡Tu propiedad se vendió!';
  else if (obj.Estado === 'Alquilada') estadoExplicado = '¡Tu propiedad se alquiló!';
  else if (obj.Estado === 'Pausada') estadoExplicado = 'Tu propiedad está pausada — contáctanos si quieres reactivarla.';
  else if (obj.Agente_Asignado_ID) estadoExplicado = 'Tu agente asignado se pondrá en contacto contigo pronto.';
  else if (obj.Modo_Venta === 'Yo solo (sin agente)') estadoExplicado = 'Estás gestionando tu propiedad de forma independiente.';

  let agenteLinkWhatsApp = '', agenteNombre = '';
  if (obj.Agente_Asignado_ID) {
    const filaAgente = buscarFilaPorId_(SHEETS.AGENTES, obj.Agente_Asignado_ID);
    if (filaAgente > -1) {
      const agenteObj = filaAObjeto_(SHEETS.AGENTES, sh_(SHEETS.AGENTES).getRange(filaAgente, 1, 1, HEADERS[SHEETS.AGENTES].length).getValues()[0]);
      agenteNombre = agenteObj.Nombre + ' ' + agenteObj.Apellido;
      agenteLinkWhatsApp = linkWhatsApp_(agenteObj.Telefono, 'Hola ' + agenteObj.Nombre + ', soy ' + obj.Nombre_Propietario + ', tengo una consulta sobre mi propiedad.');
    }
  }

  const resumenCompatibles = obtenerResumenCompradoresCompatibles_(obj);
  const analisisPropiedad = generarAnalisisPropiedad_(obj, resumenCompatibles);

  // Conexiones ya aceptadas (Conectados) de esta propiedad — para poder
  // reportar la venta desde el Portal. Solo se muestra el estado, nunca
  // el nombre del comprador (ya lo tiene por correo si fue aceptada).
  let conexionesActivas = [];
  let solicitudesDeCompradorPendientes = [];
  try {
    const sheetConexiones = sh_(SHEETS.CONEXIONES_DIRECTAS);
    const lastRowConexiones = sheetConexiones.getLastRow();
    if (lastRowConexiones > 1) {
      const todasLasFilas = sheetConexiones.getRange(2, 1, lastRowConexiones - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()
        .map(f => filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f))
        .filter(c => c.ID_Propiedad === obj.ID_Propiedad);
      conexionesActivas = todasLasFilas
        .filter(c => c.Estado === 'Conectados')
        .map(c => {
          const filaComprador = buscarFilaPorId_(SHEETS.COMPRADORES, c.ID_Comprador);
          const compradorObj = filaComprador > -1
            ? filaAObjeto_(SHEETS.COMPRADORES, sh_(SHEETS.COMPRADORES).getRange(filaComprador, 1, 1, HEADERS[SHEETS.COMPRADORES].length).getValues()[0])
            : null;
          return {
            idConexion: c.ID_Conexion,
            visitaRegistrada: !!c.Fecha_Visita_Comprador,
            linkWhatsAppComprador: (compradorObj && compradorObj.Telefono)
              ? linkWhatsApp_(compradorObj.Telefono, 'Hola ' + compradorObj.Nombre + ', soy el propietario de la vivienda que viste en PropMatchRD.')
              : ''
          };
        });
      solicitudesDeCompradorPendientes = todasLasFilas
        .filter(c => c.Estado === 'Pendiente_Respuesta_Propietario')
        .map(c => ({ idConexion: c.ID_Conexion, fecha: Utilities.formatDate(new Date(c.Fecha_Solicitud), Session.getScriptTimeZone(), 'dd/MM/yyyy') }));
    }
  } catch (e) { /* hoja aún no existe — sin conexiones que mostrar */ }

  return {
    error: false,
    idPropiedad: obj.ID_Propiedad,
    nombre: obj.Nombre_Propietario,
    operacion: obj.Operacion,
    tipoInmueble: obj.Tipo_Inmueble,
    provincia: obj.Provincia,
    municipio: obj.Municipio,
    precio: Number(obj.Precio) || 0,
    habitaciones: obj.Habitaciones,
    banos: obj.Banos,
    parqueos: obj.Parqueos,
    descripcion: obj.Descripcion,
    modoVenta: obj.Modo_Venta,
    estado: obj.Estado,
    estadoExplicado: estadoExplicado,
    verificado: obj.Verificado === 'Sí',
    agenteAsignado: agenteNombre,
    agenteLinkWhatsApp: agenteLinkWhatsApp,
    compradoresCompatiblesTotal: resumenCompatibles.total,
    compradoresCompatiblesPorClasificacion: resumenCompatibles.porClasificacion,
    analisisPropiedad: analisisPropiedad,
    conexionesActivas: conexionesActivas,
    solicitudesDeCompradorPendientes: solicitudesDeCompradorPendientes,
    urlFormularioEmbajador: PropertiesService.getScriptProperties().getProperty('FORM_EMBAJADOR_URL') || '',
    pctComisionConexionDirecta: Number(getConfig_('PCT_COMISION_CONEXION_DIRECTA', 4))
  };
}

/**
 * Instant Match — SENTIDO INVERSO: cuando se registra un comprador o
 * inquilino nuevo, busca de inmediato propiedades YA registradas que
 * sean compatibles, y notifica al comprador/inquilino. Completa el
 * Match Automático de verdad "en ambas direcciones" — antes solo
 * funcionaba cuando entraba una propiedad nueva, nunca al revés.
 */
function buscarPropiedadesCompatiblesConComprador_(comprador, esAlquiler) {
  const compatibles = [];
  const operacionBuscada = esAlquiler ? 'Alquilar' : 'Vender';
  const campoPrecioMin = esAlquiler ? 'Presupuesto_Min_Mensual' : 'Presupuesto_Min';
  const campoPrecioMax = esAlquiler ? 'Presupuesto_Max_Mensual' : 'Presupuesto_Max';

  let sheet;
  try { sheet = sh_(SHEETS.PROPIEDADES); } catch (e) { return compatibles; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return compatibles;

  const presupuestoMin = Number(comprador[campoPrecioMin]) || 0;
  const presupuestoMax = Number(comprador[campoPrecioMax]) || Infinity;

  sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.PROPIEDADES].length).getValues().forEach(f => {
    const obj = filaAObjeto_(SHEETS.PROPIEDADES, f);
    if (obj.Estado !== 'Activa') return;
    if (obj.Operacion !== operacionBuscada) return;
    // Coincide con la zona principal O con la 2da zona de interés (si tiene)
    const coincideZona = obj.Provincia === comprador.Provincia || (comprador.Provincia_2 && obj.Provincia === comprador.Provincia_2);
    if (!coincideZona) return;
    if (obj.Tipo_Inmueble && comprador.Tipo_Inmueble && obj.Tipo_Inmueble !== comprador.Tipo_Inmueble) return;
    if (Number(obj.Precio) < presupuestoMin * 0.9 || Number(obj.Precio) > presupuestoMax * 1.1) return;
    compatibles.push(obj);
  });

  return compatibles;
}

/**
 * Se llama justo después de registrar un comprador o solicitud de
 * alquiler nueva — si hay propiedades ya registradas que encajan, se le
 * avisa de inmediato en su correo de confirmación (no espera a que un
 * agente se las muestre manualmente).
 */
function obtenerBloqueInstantMatchParaComprador_(comprador, esAlquiler) {
  const compatibles = buscarPropiedadesCompatiblesConComprador_(comprador, esAlquiler);
  if (!compatibles.length) return '';

  bitacora_('Instant Match', compatibles.length + ' propiedad(es) compatible(s) encontradas para comprador nuevo', comprador.ID_Comprador || comprador.ID_Solicitud || '', '');

  const listaHtml = compatibles.slice(0, 5).map(p =>
    '<li>' + p.Tipo_Inmueble + ' en ' + p.Provincia + ' - ' + p.Municipio + ' — RD$' + Number(p.Precio).toLocaleString() + '</li>'
  ).join('');
  return '<div style="background:#f0f7ff;padding:14px;border-radius:6px;margin:10px 0">' +
    '<p style="margin:0 0 8px 0;font-weight:bold">⚡ Ya hay ' + compatibles.length + ' propiedad(es) que podrían encajar contigo:</p>' +
    '<ul style="margin:0">' + listaHtml + '</ul>' +
    (compatibles.length > 5 ? '<p style="margin:8px 0 0 0;font-size:13px">… y ' + (compatibles.length - 5) + ' más. Tu agente te las mostrará en detalle.</p>' : '') +
    '</div>';
}

/**
 * ============================================================================
 *  CONEXIÓN DIRECTA — nueva fuente de ingresos: cuando un propietario
 *  elige "Yo solo" (sin agente), puede pedir que el sistema le facilite
 *  el contacto con compradores/inquilinos compatibles YA registrados —
 *  a cambio de una comisión (4% por defecto) sobre el precio final de
 *  venta, si se concreta a través de esta conexión.
 *
 *  Diseño de privacidad: NUNCA se revela el contacto de nadie sin su
 *  consentimiento explícito — el comprador recibe primero la pregunta
 *  "¿te interesa conectar?" y solo si dice que sí, ambos lados reciben
 *  los datos de contacto del otro.
 * ============================================================================
 */
function solicitarConexionDirecta_(idPropiedad) {
  const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, idPropiedad);
  if (filaPropiedad === -1) return { error: true, mensaje: 'No encontramos tu propiedad.' };
  const propiedadesSheet = sh_(SHEETS.PROPIEDADES);
  const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, propiedadesSheet.getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);

  if (propiedad.Modo_Venta !== 'Yo solo (sin agente)') {
    return { error: true, mensaje: 'Esta opción es solo para propietarios gestionando su propiedad sin agente.' };
  }
  if (propietarioTieneComisionVencida_(propiedad.Telefono_Normalizado)) {
    return { error: true, mensaje: 'Tienes una comisión de Conexión Directa vencida sin pagar en otra propiedad. Regularízala primero desde ese Portal para poder solicitar nuevas conexiones.' };
  }

  const compatibles = buscarCompradoresCompatiblesConPropiedad_(propiedad);
  if (!compatibles.length) {
    return { error: false, mensaje: 'Por ahora no hay compradores/inquilinos compatibles registrados. Te avisaremos apenas aparezca uno.', enviadas: 0 };
  }

  const conexionesSheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const lastRow = conexionesSheet.getLastRow();
  const yaContactados = new Set();
  if (lastRow > 1) {
    conexionesSheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues().forEach(f => {
      const c = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f);
      if (c.ID_Propiedad === idPropiedad) yaContactados.add(c.ID_Comprador);
    });
  }

  const linkPortalPropietario = construirLinkPortalPropietario_(propiedad.Codigo_Portal_Privado);
  let enviadas = 0;

  compatibles.forEach(comprador => {
    if (yaContactados.has(comprador.id)) return; // ya se le preguntó antes por esta misma propiedad

    const idConexion = generarIdSecuencial_('CONEX', SHEETS.CONEXIONES_DIRECTAS, 1);
    conexionesSheet.appendRow([
      idConexion, idPropiedad, comprador.id, new Date(), 'Propietario',
      'Pendiente_Respuesta_Comprador', '', '', '', '', '', '', '', '', ''
    ]);
    enviadas++;

    // Buscar el email del comprador/inquilino para preguntarle (respeta
    // consentimiento — nunca se revela el contacto del propietario todavía)
    const esVenta = propiedad.Operacion === 'Vender';
    const sheetKey = esVenta ? SHEETS.COMPRADORES : SHEETS.SOLICITUDES_ALQUILER;
    const filaComprador = buscarFilaPorId_(sheetKey, comprador.id);
    if (filaComprador === -1) return;
    const compradorObj = filaAObjeto_(sheetKey, sh_(sheetKey).getRange(filaComprador, 1, 1, HEADERS[sheetKey].length).getValues()[0]);
    if (!compradorObj.Email) return;

    const linkPortalComprador = esVenta
      ? construirLinkPortalResultados_(compradorObj.Codigo_Portal_Privado)
      : construirLinkPortalAlquiler_(compradorObj.Codigo_Portal_Privado);

    enviarCorreo_(compradorObj.Email, '¿Te interesa esta propiedad? — ' + idPropiedad,
      '<div style="font-family:Arial"><h2>Un propietario quiere conectar contigo</h2>' +
      '<p>Un propietario que está ' + (esVenta ? 'vendiendo' : 'alquilando') + ' directamente (sin agente) tiene una propiedad que podría encajar contigo:</p>' +
      '<p><b>' + propiedad.Tipo_Inmueble + '</b> en ' + propiedad.Provincia + ' - ' + propiedad.Municipio + ' — RD$' + Number(propiedad.Precio).toLocaleString() + '</p>' +
      '<p>Si te interesa, podemos conectarte directamente con el propietario — nunca compartimos tu contacto sin que tú lo autorices primero.</p>' +
      (linkPortalComprador ? '<p style="text-align:center"><a href="' + linkPortalComprador + '" style="display:inline-block;background:#1a73e8;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Ver y responder en mi Portal</a></p>' : '') +
      '</div>');
  });

  bitacora_('Conexión Directa', 'Solicitud enviada a compradores compatibles', idPropiedad, enviadas + ' comprador(es) contactado(s)');
  return { error: false, mensaje: enviadas > 0 ? 'Le preguntamos a ' + enviadas + ' comprador(es)/inquilino(s) si les interesa conectar contigo. Te avisaremos apenas alguno responda.' : 'Ya le habíamos preguntado a todos los compradores compatibles disponibles — sin respuestas nuevas por ahora.', enviadas: enviadas };
}

/**
 * El comprador/inquilino responde si quiere conectar — solo si dice que
 * sí se revela el contacto de AMBOS lados (antes, ninguno lo tenía).
 */
function responderConexionDirecta_(idConexion, respuesta, idCompradorQueResponde) {
  const filaConexion = buscarFilaPorId_(SHEETS.CONEXIONES_DIRECTAS, idConexion);
  if (filaConexion === -1) return { error: true, mensaje: 'No encontramos esa solicitud de conexión.' };

  const sheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const conexion = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, sheet.getRange(filaConexion, 1, 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()[0]);

  // Seguridad: solo el comprador dueño de esta conexión puede responderla
  if (conexion.ID_Comprador !== idCompradorQueResponde) {
    return { error: true, mensaje: 'Esta solicitud no te pertenece.' };
  }
  if (conexion.Estado !== 'Pendiente_Respuesta_Comprador') {
    return { error: true, mensaje: 'Ya respondiste esta solicitud anteriormente.' };
  }

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Fecha_Respuesta_Comprador')).setValue(new Date());

  if (respuesta !== 'Sí') {
    sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado')).setValue('Rechazada_Por_Comprador');
    bitacora_('Conexión Directa', 'Comprador rechazó la conexión', idConexion, '');
    return { error: false, mensaje: 'Entendido, no compartiremos tu contacto con este propietario.' };
  }

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado')).setValue('Conectados');
  bitacora_('Conexión Directa', 'Comprador aceptó — contacto revelado a ambos lados', idConexion, '');

  const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, conexion.ID_Propiedad);
  const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);

  const esVenta = propiedad.Operacion === 'Vender';
  const sheetKeyComprador = esVenta ? SHEETS.COMPRADORES : SHEETS.SOLICITUDES_ALQUILER;
  const filaComprador = buscarFilaPorId_(sheetKeyComprador, conexion.ID_Comprador);
  const compradorObj = filaAObjeto_(sheetKeyComprador, sh_(sheetKeyComprador).getRange(filaComprador, 1, 1, HEADERS[sheetKeyComprador].length).getValues()[0]);

  // Revelar al PROPIETARIO el contacto del comprador
  if (propiedad.Email) {
    const linkWaComprador = linkWhatsApp_(compradorObj.Telefono, 'Hola ' + compradorObj.Nombre + ', soy el propietario de la vivienda que viste en PropMatchRD.');
    enviarCorreo_(propiedad.Email, '✅ ¡Un comprador quiere conectar contigo! — ' + conexion.ID_Propiedad,
      '<div style="font-family:Arial"><h2>✅ ' + compradorObj.Nombre + ' quiere conectar contigo</h2>' +
      '<p>Nombre: <b>' + compradorObj.Nombre + ' ' + compradorObj.Apellido + '</b> · Teléfono: <b>' + compradorObj.Telefono + '</b></p>' +
      '<p><a href="' + linkWaComprador + '" style="background:#25D366;color:white;padding:10px 16px;border-radius:6px;text-decoration:none">Contactar por WhatsApp</a></p>' +
      '<p style="font-size:13px;color:#666">Recuerda: si la venta se concreta a través de esta conexión, aplica una comisión de ' + getConfig_('PCT_COMISION_CONEXION_DIRECTA', 4) + '% sobre el precio final — repórtala desde tu Portal cuando se cierre.</p>' +
      '</div>');
  }
  // Revelar al COMPRADOR el contacto del propietario
  if (compradorObj.Email) {
    const linkWaPropietario = linkWhatsApp_(propiedad.Telefono, 'Hola ' + propiedad.Nombre_Propietario + ', vi tu propiedad en PropMatchRD y me interesa.');
    enviarCorreo_(compradorObj.Email, '✅ ¡El propietario aceptó conectar! — ' + conexion.ID_Propiedad,
      '<div style="font-family:Arial"><h2>✅ Ya puedes contactar al propietario</h2>' +
      '<p>Nombre: <b>' + propiedad.Nombre_Propietario + ' ' + propiedad.Apellido_Propietario + '</b> · Teléfono: <b>' + propiedad.Telefono + '</b></p>' +
      '<p><a href="' + linkWaPropietario + '" style="background:#25D366;color:white;padding:10px 16px;border-radius:6px;text-decoration:none">Contactar por WhatsApp</a></p>' +
      '</div>');
  }

  return { error: false, mensaje: '¡Listo! Te enviamos por correo los datos de contacto del propietario.' };
}

/**
 * El propietario reporta que la venta se concretó con un comprador
 * conectado a través del sistema — se calcula la comisión de la
 * plataforma (4% por defecto) y se le muestra la cuenta bancaria para
 * pagarla, mismo mecanismo ya usado para comisiones de agentes.
 */
function reportarVentaConexionDirecta_(idPropiedad, idConexion, montoVenta) {
  const filaConexion = buscarFilaPorId_(SHEETS.CONEXIONES_DIRECTAS, idConexion);
  if (filaConexion === -1) return { error: true, mensaje: 'No encontramos esa conexión.' };
  const sheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const conexion = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, sheet.getRange(filaConexion, 1, 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()[0]);

  if (conexion.ID_Propiedad !== idPropiedad) return { error: true, mensaje: 'Esta conexión no corresponde a tu propiedad.' };
  if (conexion.Estado !== 'Conectados') return { error: true, mensaje: 'Esta conexión todavía no fue aceptada por el comprador.' };

  const monto = Number(montoVenta) || 0;
  if (monto <= 0) return { error: true, mensaje: 'Ingresa un monto de venta válido.' };

  const pct = Number(getConfig_('PCT_COMISION_CONEXION_DIRECTA', 4));
  const comisionRD = Math.round(monto * (pct / 100));
  const diasLimite = Number(getConfig_('DIAS_LIMITE_PAGO_CONEXION_DIRECTA', 15));
  const fechaLimite = new Date();
  fechaLimite.setDate(fechaLimite.getDate() + diasLimite);

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado')).setValue('Venta_Concretada');
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Fecha_Venta_Concretada')).setValue(new Date());
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Monto_Venta_RD')).setValue(monto);
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Comision_Plataforma_Pct')).setValue(pct);
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Comision_Plataforma_RD')).setValue(comisionRD);
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado_Pago_Comision')).setValue('Pendiente');
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Fecha_Limite_Pago')).setValue(fechaLimite);

  // Marcar la propiedad como vendida/alquilada
  const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, idPropiedad);
  const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);
  sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, col_(SHEETS.PROPIEDADES, 'Estado')).setValue(propiedad.Operacion === 'Vender' ? 'Vendida' : 'Alquilada');

  bitacora_('Conexión Directa', 'Venta concretada, comisión calculada', idConexion, 'RD$' + comisionRD.toLocaleString() + ' (' + pct + '%)');

  const emailAdmin = getConfig_('EMAIL_ADMIN', '');
  if (emailAdmin) {
    enviarCorreo_(emailAdmin, '💰 Comisión de conexión directa por cobrar — ' + idConexion,
      '<p>Propiedad <b>' + idPropiedad + '</b> vendida/alquilada por conexión directa (sin agente). Monto: RD$' + monto.toLocaleString() + '. Comisión (' + pct + '%): <b>RD$' + comisionRD.toLocaleString() + '</b>.</p>');
  }

  return {
    error: false,
    mensaje: '¡Felicidades por tu venta! Aplica una comisión de RD$' + comisionRD.toLocaleString() + ' (' + pct + '%) para la plataforma.',
    comisionRD: comisionRD,
    cuentaPlataforma: obtenerCuentaPlataformaFormateada_(),
    referenciaPago: idConexion,
    urlQrPago: getConfig_('URL_QR_PAGO_AZUL', '')
  };
}

/** El propietario reporta que ya transfirió la comisión de una conexión directa */
function reportarPagoConexionDirecta_(idConexion) {
  const filaConexion = buscarFilaPorId_(SHEETS.CONEXIONES_DIRECTAS, idConexion);
  if (filaConexion === -1) return { error: true, mensaje: 'No encontramos esa conexión.' };
  const sheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado_Pago_Comision')).setValue('Pago reportado por propietario');
  bitacora_('Conexión Directa', 'Propietario reportó pago de comisión', idConexion, '');

  const emailAdmin = getConfig_('EMAIL_ADMIN', '');
  if (emailAdmin) {
    enviarCorreo_(emailAdmin, '💰 Propietario reportó pago — ' + idConexion,
      '<p>Verifica que el pago llegó y marca "Pagado" en la hoja "Conexiones Directas" para confirmarle al propietario.</p>');
  }
  return { error: false, mensaje: 'Gracias, verificaremos tu pago y te confirmaremos pronto.' };
}

/** Solicitudes de conexión directa pendientes de respuesta para un comprador/inquilino específico */
function obtenerConexionesDirectasPendientes_(idComprador) {
  let sheet;
  try { sheet = sh_(SHEETS.CONEXIONES_DIRECTAS); } catch (e) { return []; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const pendientes = [];
  sheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues().forEach(f => {
    const c = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f);
    if (c.ID_Comprador !== idComprador || c.Estado !== 'Pendiente_Respuesta_Comprador') return;
    const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, c.ID_Propiedad);
    if (filaPropiedad === -1) return;
    const p = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);
    pendientes.push({
      idConexion: c.ID_Conexion, tipoInmueble: p.Tipo_Inmueble, provincia: p.Provincia,
      municipio: p.Municipio, precio: Number(p.Precio) || 0, operacion: p.Operacion
    });
  });
  return pendientes;
}

/**
 * ============================================================================
 *  SEGURIDAD DE COBRO — Conexión Directa (mismo respaldo que ya existe
 *  para agentes, aplicado ahora a propietarios en modo "Yo solo")
 * ============================================================================
 *  Antes, un propietario podía dejar de pagar su comisión de Conexión
 *  Directa sin ninguna consecuencia — a diferencia de los agentes, que sí
 *  quedan bloqueados automáticamente si tienen una comisión vencida. Esto
 *  cierra ese hueco: marca automáticamente las comisiones vencidas y
 *  bloquea nuevas solicitudes de conexión hasta que se regularice.
 * ============================================================================
 */
function sincronizarEstadoConexionesDirectas_() {
  let sheet;
  try { sheet = sh_(SHEETS.CONEXIONES_DIRECTAS); } catch (e) { return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = HEADERS[SHEETS.CONEXIONES_DIRECTAS];
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const colEstadoPago = col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado_Pago_Comision');
  const hoy = new Date();

  data.forEach((fila, i) => {
    const obj = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, fila);
    if (obj.Estado_Pago_Comision === 'Pendiente' && obj.Fecha_Limite_Pago && new Date(obj.Fecha_Limite_Pago) < hoy) {
      sheet.getRange(i + 2, colEstadoPago).setValue('Vencido');
      bitacora_('Sistema', 'Comisión de Conexión Directa vencida', obj.ID_Conexion, 'Propiedad ' + obj.ID_Propiedad);

      const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, obj.ID_Propiedad);
      if (filaPropiedad > -1) {
        const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);
        if (propiedad.Email) {
          enviarCorreo_(propiedad.Email, '⚠️ Comisión vencida — ' + obj.ID_Propiedad,
            '<p>La comisión de RD$' + Number(obj.Comision_Plataforma_RD || 0).toLocaleString() + ' por tu conexión directa venció sin pago. ' +
            'No podrás solicitar nuevas conexiones hasta regularizarla — repórtala desde tu Portal en cuanto la transfieras.</p>');
        }
      }
    }
  });
}

/**
 * Verifica si un propietario (por teléfono, para cubrir todas sus
 * propiedades) tiene alguna comisión de Conexión Directa vencida sin
 * pagar — si es así, no se le permite solicitar nuevas conexiones hasta
 * regularizar.
 */
function propietarioTieneComisionVencida_(telefonoNormalizado) {
  if (!telefonoNormalizado) return false;
  let sheetConexiones, sheetPropiedades;
  try {
    sheetConexiones = sh_(SHEETS.CONEXIONES_DIRECTAS);
    sheetPropiedades = sh_(SHEETS.PROPIEDADES);
  } catch (e) { return false; }

  const lastRowConexiones = sheetConexiones.getLastRow();
  if (lastRowConexiones < 2) return false;

  const idsPropiedadesDelMismoTelefono = new Set();
  const lastRowProp = sheetPropiedades.getLastRow();
  if (lastRowProp > 1) {
    sheetPropiedades.getRange(2, 1, lastRowProp - 1, HEADERS[SHEETS.PROPIEDADES].length).getValues().forEach(f => {
      const p = filaAObjeto_(SHEETS.PROPIEDADES, f);
      if (p.Telefono_Normalizado === telefonoNormalizado) idsPropiedadesDelMismoTelefono.add(p.ID_Propiedad);
    });
  }

  return sheetConexiones.getRange(2, 1, lastRowConexiones - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()
    .some(f => {
      const c = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f);
      return idsPropiedadesDelMismoTelefono.has(c.ID_Propiedad) && c.Estado_Pago_Comision === 'Vencido';
    });
}

/**
 * ============================================================================
 *  CONEXIÓN DIRECTA INICIADA POR EL COMPRADOR/INQUILINO
 * ============================================================================
 *  Espejo de solicitarConexionDirecta_ (que la inicia el propietario) —
 *  aquí es el COMPRADOR quien ve una propiedad compatible en su Portal y
 *  pide conectar. El propietario es quien debe aceptar/rechazar (mismo
 *  principio de consentimiento mutuo: nadie ve el contacto del otro hasta
 *  que AMBOS acepten).
 * ============================================================================
 */
function solicitarConexionDesdeComprador_(idComprador, idPropiedad, esAlquiler) {
  const sheetKeyComprador = esAlquiler ? SHEETS.SOLICITUDES_ALQUILER : SHEETS.COMPRADORES;
  const filaComprador = buscarFilaPorId_(sheetKeyComprador, idComprador);
  if (filaComprador === -1) return { error: true, mensaje: 'No encontramos tu registro.' };
  const compradorObj = filaAObjeto_(sheetKeyComprador, sh_(sheetKeyComprador).getRange(filaComprador, 1, 1, HEADERS[sheetKeyComprador].length).getValues()[0]);

  const filaPropiedad = buscarFilaPorId_(SHEETS.PROPIEDADES, idPropiedad);
  if (filaPropiedad === -1) return { error: true, mensaje: 'Esa propiedad ya no está disponible.' };
  const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(filaPropiedad, 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);

  if (propiedad.Modo_Venta !== 'Yo solo (sin agente)' || propiedad.Estado !== 'Activa') {
    return { error: true, mensaje: 'Esta propiedad ya no está disponible para conexión directa.' };
  }

  const conexionesSheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const lastRow = conexionesSheet.getLastRow();
  if (lastRow > 1) {
    const yaExiste = conexionesSheet.getRange(2, 1, lastRow - 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()
      .some(f => {
        const c = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, f);
        return c.ID_Propiedad === idPropiedad && c.ID_Comprador === idComprador;
      });
    if (yaExiste) return { error: true, mensaje: 'Ya solicitaste conectar con esta propiedad anteriormente.' };
  }

  const idConexion = generarIdSecuencial_('CONEX', SHEETS.CONEXIONES_DIRECTAS, 1);
  conexionesSheet.appendRow([
    idConexion, idPropiedad, idComprador, new Date(), 'Comprador',
    'Pendiente_Respuesta_Propietario', '', '', '', '', '', '', '', '', ''
  ]);
  bitacora_('Conexión Directa', 'Comprador solicitó conectar con propiedad', idConexion, idPropiedad);

  if (propiedad.Email) {
    const linkPortalPropietario = construirLinkPortalPropietario_(propiedad.Codigo_Portal_Privado);
    enviarCorreo_(propiedad.Email, '¿Te interesa conectar con este comprador? — ' + idPropiedad,
      '<div style="font-family:Arial"><h2>Un comprador quiere conectar contigo</h2>' +
      '<p>Alguien interesado en tu propiedad (' + propiedad.Tipo_Inmueble + ' en ' + propiedad.Provincia + ' - ' + propiedad.Municipio + ') solicitó conectar directamente.</p>' +
      '<p>Nunca compartimos su contacto sin que tú lo autorices primero.</p>' +
      (linkPortalPropietario ? '<p style="text-align:center"><a href="' + linkPortalPropietario + '" style="display:inline-block;background:#1a73e8;color:white;padding:12px 24px;border-radius:8px;text-decoration:none;font-weight:bold">Ver y responder en mi Portal</a></p>' : '') +
      '</div>');
  }

  return { error: false, mensaje: 'Solicitud enviada — te avisaremos apenas el propietario responda.' };
}

/** Espejo de responderConexionDirecta_ — aquí responde el PROPIETARIO (la solicitud la inició el comprador) */
function responderConexionDirectaPropietario_(idConexion, respuesta, idPropiedadQueResponde) {
  const filaConexion = buscarFilaPorId_(SHEETS.CONEXIONES_DIRECTAS, idConexion);
  if (filaConexion === -1) return { error: true, mensaje: 'No encontramos esa solicitud de conexión.' };

  const sheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const conexion = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, sheet.getRange(filaConexion, 1, 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()[0]);

  if (conexion.ID_Propiedad !== idPropiedadQueResponde) {
    return { error: true, mensaje: 'Esta solicitud no te pertenece.' };
  }
  if (conexion.Estado !== 'Pendiente_Respuesta_Propietario') {
    return { error: true, mensaje: 'Ya respondiste esta solicitud anteriormente.' };
  }

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Fecha_Respuesta_Comprador')).setValue(new Date());

  if (respuesta !== 'Sí') {
    sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado')).setValue('Rechazada_Por_Propietario');
    bitacora_('Conexión Directa', 'Propietario rechazó la conexión', idConexion, '');
    return { error: false, mensaje: 'Entendido, no compartiremos tu contacto con este comprador.' };
  }

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Estado')).setValue('Conectados');
  bitacora_('Conexión Directa', 'Propietario aceptó — contacto revelado a ambos lados', idConexion, '');

  const propiedad = filaAObjeto_(SHEETS.PROPIEDADES, sh_(SHEETS.PROPIEDADES).getRange(buscarFilaPorId_(SHEETS.PROPIEDADES, conexion.ID_Propiedad), 1, 1, HEADERS[SHEETS.PROPIEDADES].length).getValues()[0]);
  const esVenta = propiedad.Operacion === 'Vender';
  const sheetKeyComprador = esVenta ? SHEETS.COMPRADORES : SHEETS.SOLICITUDES_ALQUILER;
  const compradorObj = filaAObjeto_(sheetKeyComprador, sh_(sheetKeyComprador).getRange(buscarFilaPorId_(sheetKeyComprador, conexion.ID_Comprador), 1, 1, HEADERS[sheetKeyComprador].length).getValues()[0]);

  if (propiedad.Email) {
    const linkWaComprador = linkWhatsApp_(compradorObj.Telefono, 'Hola ' + compradorObj.Nombre + ', soy el propietario de la vivienda que viste en PropMatchRD.');
    enviarCorreo_(propiedad.Email, '✅ ¡Conexión aceptada! — ' + conexion.ID_Propiedad,
      '<div style="font-family:Arial"><h2>✅ Ya puedes contactar a ' + compradorObj.Nombre + '</h2>' +
      '<p>Nombre: <b>' + compradorObj.Nombre + ' ' + compradorObj.Apellido + '</b> · Teléfono: <b>' + compradorObj.Telefono + '</b></p>' +
      '<p><a href="' + linkWaComprador + '" style="background:#25D366;color:white;padding:10px 16px;border-radius:6px;text-decoration:none">Contactar por WhatsApp</a></p>' +
      '</div>');
  }
  if (compradorObj.Email) {
    const linkWaPropietario = linkWhatsApp_(propiedad.Telefono, 'Hola, vi que aceptaste mi solicitud de conexión en PropMatchRD sobre tu propiedad.');
    enviarCorreo_(compradorObj.Email, '✅ ¡El propietario aceptó conectar! — ' + conexion.ID_Propiedad,
      '<div style="font-family:Arial"><h2>✅ Ya puedes contactar al propietario</h2>' +
      '<p>Nombre: <b>' + propiedad.Nombre_Propietario + '</b> · Teléfono: <b>' + propiedad.Telefono + '</b></p>' +
      '<p><a href="' + linkWaPropietario + '" style="background:#25D366;color:white;padding:10px 16px;border-radius:6px;text-decoration:none">Contactar por WhatsApp</a></p>' +
      '</div>');
  }

  return { error: false, mensaje: '¡Contacto revelado! Ya pueden coordinar directamente.' };
}

/** El comprador marca que visitó una propiedad conectada por Conexión Directa — registro simple, sin depender del agente */
function registrarVisitaConexionDirecta_(idConexion, idCompradorQueRegistra) {
  const filaConexion = buscarFilaPorId_(SHEETS.CONEXIONES_DIRECTAS, idConexion);
  if (filaConexion === -1) return { error: true, mensaje: 'No encontramos esa conexión.' };

  const sheet = sh_(SHEETS.CONEXIONES_DIRECTAS);
  const conexion = filaAObjeto_(SHEETS.CONEXIONES_DIRECTAS, sheet.getRange(filaConexion, 1, 1, HEADERS[SHEETS.CONEXIONES_DIRECTAS].length).getValues()[0]);

  if (conexion.ID_Comprador !== idCompradorQueRegistra) return { error: true, mensaje: 'Esta conexión no te pertenece.' };
  if (conexion.Estado !== 'Conectados') return { error: true, mensaje: 'Todavía no se ha aceptado esta conexión.' };

  sheet.getRange(filaConexion, col_(SHEETS.CONEXIONES_DIRECTAS, 'Fecha_Visita_Comprador')).setValue(new Date());
  bitacora_('Conexión Directa', 'Comprador registró que visitó la propiedad', idConexion, '');
  return { error: false, mensaje: '¡Registrado! Cuando llegues a un acuerdo, el propietario podrá reportar la venta desde su Portal.' };
}
