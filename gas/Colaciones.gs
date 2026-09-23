// COLACIONES SEMANALES - Google Apps Script backend
// Controla la inscripcion semanal al menu de colaciones de los trabajadores:
// evita que alguien quede sin anotarse y evita anotaciones duplicadas.
//
// Hojas requeridas en la planilla de Google Sheets ligada a este script:
//   Trabajadores: RUT | Nombre | Tipo | Activo | FechaInicio | FechaFin | Notas
//     - Tipo: "Fijo" o "Spot" (visitas/clientes temporales)
//     - Activo: "Si" o "No"
//     - FechaInicio / FechaFin: solo se usan para Tipo=Spot (rango de vigencia,
//       formato AAAA-MM-DD). Si estan vacias, se considera vigente siempre.
//   Menus: Semana | Dia | Opcion | Descripcion | Activo | Especial | DescripcionEspecial
//     - Semana: fecha del lunes de esa semana, formato AAAA-MM-DD
//     - Dia: Lunes, Martes, Miercoles, Jueves, Viernes
//     - Opcion: A, B, C...
//     - Especial: "Si" o "No" - marca un dia como almuerzo mejorado (para
//       destacarlo y llevar registro de quienes se anotaron ese dia)
//     - DescripcionEspecial: texto libre con lo que incluye el almuerzo
//       mejorado ese dia (ej. "Torta con bebida"). Se repite en todas las
//       filas/opciones de ese dia, igual que Especial.
//   Pedidos: ID | Semana | RUT | Nombre | Dia | Opcion | Timestamp
//     - Un trabajador solo puede tener UNA fila por (Semana, RUT, Dia): al
//       guardar un pedido para un dia ya elegido, se reemplaza la opcion en
//       vez de crear una fila nueva (evita anotaciones dobles).
//   Config: Clave | Valor
//     - admin_password: clave de acceso del panel de administracion
//     - cierre_dias_antes: cuantos dias antes del lunes de la semana se
//       cierran las inscripciones (por defecto 5 = miercoles de la semana
//       previa). Editable desde el panel de administracion.
//     - cierre_hora: hora de cierre ese dia, formato HH:MM (por defecto
//       "14:00"). Editable desde el panel de administracion. Se guarda
//       forzando formato de texto en la celda para que Sheets no la
//       reinterprete como una hora/fecha.
//     Pasado ese plazo, los trabajadores ya no pueden elegir ni cambiar su
//     opcion para esa semana: solo la administracion puede seguir editando
//     pedidos de esa semana (guardarPedido con asAdmin:true).
//   Platos: Nombre
//     - Catalogo de platos ya usados alguna vez, para elegirlos rapido al
//       armar el menu en vez de escribirlos de nuevo. Se llena solo: cada
//       vez que se guarda una opcion de menu con una descripcion nueva, se
//       agrega aqui si no existia (sin duplicados, sin distinguir mayus).
//   Semanas: Semana | Publicada
//     - Controla si los trabajadores ya pueden ver el menu de esa semana.
//       El admin puede ir guardando el menu de a poco (queda en borrador,
//       "Publicada"="No") y solo cuando aprieta "Publicar semana" los
//       trabajadores lo ven. Una semana sin fila aqui se considera
//       publicada (compatibilidad con semanas armadas antes de este
//       control). La primera opcion que se guarda para una semana nueva la
//       deja automaticamente en borrador.
//   Feriados: Semana | Dia
//     - Marca un dia especifico de una semana como feriado (no se trabaja
//       ese dia): los trabajadores dejan de ver/poder elegir ese dia y no se
//       cuenta colacion. Al marcar un dia como feriado se borran los
//       pedidos que ya existieran ahi, porque ese dia no habra colacion.

const HOJAS_COL = {
  TRABAJADORES: 'Trabajadores',
  MENUS: 'Menus',
  PEDIDOS: 'Pedidos',
  CONFIG: 'Config',
  PLATOS: 'Platos',
  SEMANAS: 'Semanas',
  FERIADOS: 'Feriados'
};

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const bodyParam = e && e.parameter && e.parameter.body;

  if (bodyParam) {
    try {
      const body = JSON.parse(decodeURIComponent(bodyParam));
      return responderJsonpCol(procesarAccionCol(body), callback);
    } catch (err) {
      return responderJsonpCol({ ok: false, error: err.toString() }, callback);
    }
  }

  try {
    return responderJsonpCol(obtenerTodoCol(), callback);
  } catch (err) {
    return responderJsonpCol({ ok: false, error: err.toString() }, callback);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return responderCol(procesarAccionCol(body));
  } catch (err) {
    return responderCol({ ok: false, error: err.toString(), stack: err.stack });
  }
}

// Acciones que solo administracion puede ejecutar. Antes, cualquiera que
// descubriera la URL /exec (visible igual en el codigo del navegador, con
// "Ver codigo fuente") podia llamarlas directamente sin conocer la clave de
// admin: crear/borrar trabajadores, editar el menu, cambiar la clave, etc.
// Ahora cada una exige la clave vigente en Config junto con la peticion.
const ACCIONES_SOLO_ADMIN_COL = [
  'guardarTrabajador', 'eliminarTrabajador', 'guardarMenu', 'eliminarMenu',
  'marcarDiaEspecial', 'marcarSemanaPublicada', 'marcarFeriado', 'guardarConfig', 'eliminarPlato'
];
function claveAdminValidaCol(password) {
  const clave = String(configCacheadoCol().admin_password || '').trim();
  return !!clave && String(password || '').trim() === clave;
}
function procesarAccionCol(body) {
  const requiereAdmin = ACCIONES_SOLO_ADMIN_COL.indexOf(body.action) !== -1 ||
    ((body.action === 'guardarPedido' || body.action === 'eliminarPedido') && body.asAdmin);
  if (requiereAdmin && !claveAdminValidaCol(body.adminPassword)) {
    return { ok: false, error: 'No autorizado. Vuelve a iniciar sesion como administracion.' };
  }
  switch (body.action) {
    case 'loginAdmin':
      return verificarLoginAdminCol(body.password);
    case 'loginTrabajador':
      return loginTrabajadorCol(body.rut);
    case 'guardarTrabajador':
      return guardarTrabajadorCol(body.data, body.isEdit);
    case 'eliminarTrabajador':
      return eliminarTrabajadorCol(body.rut);
    case 'guardarMenu':
      return guardarMenuCol(body.data);
    case 'eliminarMenu':
      return eliminarMenuCol(body.semana, body.dia, body.opcion);
    case 'marcarDiaEspecial':
      return marcarDiaEspecialCol(body.semana, body.dia, body.especial, body.descripcionEspecial);
    case 'marcarSemanaPublicada':
      return marcarSemanaPublicadaCol(body.semana, body.publicada);
    case 'marcarFeriado':
      return marcarFeriadoCol(body.semana, body.dia, body.feriado);
    case 'guardarPedido':
      return guardarPedidoCol(body.data);
    case 'eliminarPedido':
      return eliminarPedidoCol(body.semana, body.rut, body.dia, body.asAdmin);
    case 'guardarConfig':
      return actualizarConfigCol(body.clave, body.valor);
    case 'eliminarPlato':
      return eliminarPlatoCol(body.nombre);
    default:
      return { ok: false, error: 'Accion no reconocida: ' + body.action };
  }
}

// --- RESPUESTAS ---
function responderCol(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
function responderJsonpCol(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// --- CACHE (para que 50+ personas entrando a la vez no golpeen todas la
// planilla al mismo tiempo) ---
// Cada login/actualizacion antes hacia 3-4 lecturas completas de hojas
// (Trabajadores, Menus, Config, Pedidos), y ese costo se repetia por cada
// persona conectada. CacheService comparte estos resultados por unos
// segundos entre TODAS las ejecuciones del script, asi que si 50 personas
// entran en la misma ventana de tiempo, solo la primera paga el costo de
// leer la hoja; el resto reutiliza el mismo resultado cacheado. Los
// pedidos se cachean muy poco (5s) porque cambian todo el tiempo; el
// resto un poco mas porque cambia rara vez.
// `esValido` (opcional) evita cachear un resultado "vacio por error": por
// ejemplo, si leer Config falla transitoriamente (posible bajo carga con
// 60 personas a la vez) y devuelve {} en vez de lanzar, sin este chequeo
// ese {} quedaria cacheado 20s y durante ese rato toda accion de admin
// fallaria como "No autorizado" y el cierre de inscripciones usaria los
// valores por defecto en vez de los configurados.
function cacheColLeer_(clave, ttlSeg, fn, esValido) {
  const cache = CacheService.getScriptCache();
  try {
    const guardado = cache.get(clave);
    if (guardado != null) return JSON.parse(guardado);
  } catch (err) { /* cache ilegible: seguir y recalcular */ }
  const valor = fn();
  if (!esValido || esValido(valor)) {
    try { cache.put(clave, JSON.stringify(valor), ttlSeg); } catch (err) { /* valor muy grande para cachear: no es grave */ }
  }
  return valor;
}
function cacheColInvalidar_(clave) {
  try { CacheService.getScriptCache().remove(clave); } catch (err) { /* nada que invalidar */ }
}
function trabajadoresCacheadosCol() { return cacheColLeer_('trabajadores', 20, () => hojaAObjetosCol(HOJAS_COL.TRABAJADORES)); }
function menusCacheadosCol() { return cacheColLeer_('menus', 15, menusAObjetosCol); }
function platosCacheadosCol() { return cacheColLeer_('platos', 30, listarPlatosCol); }
function configCacheadoCol() { return cacheColLeer_('config', 20, obtenerConfigCol, v => v && Object.keys(v).length > 0); }
// TTL subido de 5s a 20s: como cacheColInvalidar_('pedidos') se llama justo
// despues de CADA guardarPedido/eliminarPedido, un TTL mas largo no arriesga
// mostrar un pedido desactualizado tras escribir (se invalida al toque) —
// solo evita que, con muchas personas leyendo a la vez (por ejemplo varios
// que inician sesion casi al mismo tiempo), cada una dispare su propia
// lectura completa de la hoja Pedidos en vez de reusar la misma cache.
function pedidosCacheadosCol() { return cacheColLeer_('pedidos', 20, () => hojaAObjetosCol(HOJAS_COL.PEDIDOS)); }
function feriadosCacheadosCol() {
  return cacheColLeer_('feriados', 20, () => {
    try { return hojaAObjetosCol(HOJAS_COL.FERIADOS); } catch (err) { return []; }
  });
}
// Un solo candado para todo el script: evita que dos personas escribiendo
// pedidos/menus al mismo tiempo se pisen (por ejemplo, dos filas para el
// mismo dia y persona si sus peticiones se entrelazan). Con 50 personas
// escribiendo casi a la vez esto es lo que evita datos duplicados o
// inconsistentes, a cambio de una espera muy breve si coinciden.
function conCandadoCol_(fn) {
  const candado = LockService.getScriptLock();
  try {
    candado.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'El servidor está muy ocupado en este momento, intenta de nuevo en unos segundos.' };
  }
  try {
    return fn();
  } finally {
    candado.releaseLock();
  }
}

// --- HELPERS ---
// Cachea la planilla activa durante esta ejecucion: una peticion tipica
// (ej. obtenerTodoCol) toca hasta 7 hojas distintas, y cada llamada a
// SpreadsheetApp.getActiveSpreadsheet() tiene su propio costo. Sin esto se
// pagaba ese costo una vez por hoja en vez de una sola vez por peticion,
// lo que se sumaba a la lentitud real observada en "Ejecuciones" (varios
// doGet/doPost de 5-12 segundos).
var _ssActivaCol = null;
function ssActivaCol_() {
  if (!_ssActivaCol) _ssActivaCol = SpreadsheetApp.getActiveSpreadsheet();
  return _ssActivaCol;
}
function getHojaCol(nombre) {
  const ss = ssActivaCol_();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    const nombreLower = String(nombre).toLowerCase();
    h = ss.getSheets().find(s => s.getName().toLowerCase() === nombreLower);
  }
  if (!h) throw new Error('Hoja no encontrada: ' + nombre);
  return h;
}
function normalizarRutCol(rut) {
  return String(rut || '').toLowerCase().replace(/\./g, '').replace(/-/g, '').replace(/\s/g, '').trim();
}
function formatearFechaCol(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').trim();
}
// Timestamp en hora local (segun el huso horario configurado en el
// proyecto de Apps Script / la planilla) en vez de UTC, para que la
// columna Timestamp de Pedidos coincida con la hora real de Chile.
// Si el huso horario del script esta mal configurado, este timestamp
// seguira desfasado: revisar Configuracion del proyecto (icono de
// engranaje) -> Zona horaria -> America/Santiago.
function timestampLocalCol() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}
// Compara una celda (que Sheets puede haber convertido a Date) contra un
// texto plano tipo 'AAAA-MM-DD', normalizando ambos lados igual.
function mismaFechaCol(celda, texto) {
  return formatearFechaCol(celda) === formatearFechaCol(texto);
}
function hojaAObjetosCol(nombreHoja) {
  const h = getHojaCol(nombreHoja);
  const v = h.getDataRange().getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(c => String(c).trim().toLowerCase());
  return v.slice(1)
    .filter(f => f.some(c => String(c).trim() !== ''))
    .map((f, idx) => {
      const obj = { _fila: idx + 2 };
      cab.forEach((c, i) => {
        let val = f[i];
        if (val instanceof Date) val = formatearFechaCol(val);
        obj[c] = val != null ? val : '';
      });
      return obj;
    });
}
// Lectura de Menus por posicion de columna (no por el texto de la
// cabecera): si a alguien se le corrompio o le falta el titulo de la
// columna G (DescripcionEspecial) -por ejemplo en una planilla creada antes
// de agregar esa columna y donde nunca se corrio agregarColumnaDescripcionEspecial
// prolijamente-, hojaAObjetosCol() no generaba la clave "descripcionespecial"
// y el detalle del almuerzo mejorado se guardaba en la hoja pero nunca le
// llegaba al trabajador. Esta funcion siempre devuelve las 7 claves fijas.
function menusAObjetosCol() {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  if (v.length < 2) return [];
  return v.slice(1)
    .filter(f => f.some(c => String(c).trim() !== ''))
    .map((f, idx) => ({
      _fila: idx + 2,
      semana: f[0] instanceof Date ? formatearFechaCol(f[0]) : (f[0] != null ? f[0] : ''),
      dia: f[1] != null ? f[1] : '',
      opcion: f[2] != null ? f[2] : '',
      descripcion: f[3] != null ? f[3] : '',
      activo: f[4] != null ? f[4] : '',
      especial: f[5] != null ? f[5] : '',
      descripcionespecial: f[6] != null ? f[6] : ''
    }));
}
// Nunca se manda admin_password al navegador: obtenerTodoCol() responde a
// una peticion GET simple sin ninguna autenticacion (es el "?" o el boton
// Actualizar del admin), asi que si se filtrara ahi, cualquiera con la URL
// del script -visible igual en el codigo de la pagina- podria leer la
// clave de administrador en texto plano.
function configSinClaveCol(cfg) {
  const copia = Object.assign({}, cfg);
  delete copia.admin_password;
  return copia;
}
// A los trabajadores solo se les manda el menu de semanas ya publicadas
// (ver semanaPublicadaCol): una semana en borrador no debe ni llegar a su
// navegador, no solo estar oculta en la pantalla.
function menusPublicadosCol(menus) {
  return menus.filter(m => semanaPublicadaCol(m.semana));
}
function obtenerTodoCol() {
  return {
    ok: true,
    trabajadores: trabajadoresCacheadosCol(),
    menus: menusCacheadosCol(),
    pedidos: pedidosCacheadosCol(),
    platos: platosCacheadosCol(),
    config: configSinClaveCol(configCacheadoCol()),
    semanas: semanasCacheadasCol(),
    feriados: feriadosCacheadosCol(),
    timestamp: new Date().toISOString()
  };
}
function obtenerConfigCol() {
  try {
    const v = getHojaCol(HOJAS_COL.CONFIG).getDataRange().getValues();
    const c = {};
    v.slice(1).forEach(f => {
      if (!f[0]) return;
      let valor = f[1];
      // Si la celda quedo guardada como fecha/hora (Sheets la autodetecto
      // como tal alguna vez), se recupera solo la hora:minuto en vez de
      // devolver el timestamp completo tal cual.
      if (valor instanceof Date) valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), 'HH:mm');
      c[String(f[0]).toLowerCase().trim()] = valor;
    });
    return c;
  } catch (err) {
    return {};
  }
}
function actualizarConfigCol(clave, valor) {
  if (!clave) return { ok: false, error: 'Falta la clave de configuracion' };
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.CONFIG);
    const v = h.getDataRange().getValues();
    // Forzar formato de texto en la celda para que Sheets no reinterprete
    // valores como "14:00" como una hora/fecha (lo que corrompia cierre_hora).
    let resultado;
    let encontrado = false;
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]).toLowerCase().trim() === String(clave).toLowerCase().trim()) {
        h.getRange(i + 1, 2).setNumberFormat('@').setValue(valor);
        resultado = { ok: true };
        encontrado = true;
        break;
      }
    }
    if (!encontrado) {
      const fila = h.getLastRow() + 1;
      h.getRange(fila, 1).setValue(clave);
      h.getRange(fila, 2).setNumberFormat('@').setValue(valor);
      resultado = { ok: true, creado: true };
    }
    cacheColInvalidar_('config');
    return resultado;
  });
}
// Calcula el instante (Date) en que se cierran las inscripciones de una
// semana: "cierre_dias_antes" dias antes del lunes de esa semana, a la hora
// "cierre_hora". Por defecto: miercoles de la semana previa a las 14:00.
function calcularCierreCol(semana, cfg) {
  cfg = cfg || configCacheadoCol();
  const diasAntes = Number(cfg.cierre_dias_antes);
  const dias = isNaN(diasAntes) ? 5 : diasAntes;
  const hora = String(cfg.cierre_hora || '14:00').trim();
  const partes = hora.split(':');
  const hh = Number(partes[0]) || 0;
  const mm = Number(partes[1]) || 0;
  const partesSemana = String(semana).split('-').map(Number);
  const lunes = new Date(partesSemana[0], (partesSemana[1] || 1) - 1, partesSemana[2] || 1);
  const cierre = new Date(lunes);
  cierre.setDate(cierre.getDate() - dias);
  cierre.setHours(hh, mm, 0, 0);
  return cierre;
}
function inscripcionesCerradasCol(semana, cfg) {
  return new Date() > calcularCierreCol(semana, cfg);
}

// --- ADMIN LOGIN ---
// Devuelve, en la misma llamada, todos los datos que necesita el panel de
// administracion (evita un segundo viaje al servidor solo para cargarlos,
// que es lo que hacia mas lento el ingreso como admin).
function verificarLoginAdminCol(password) {
  const cfg = configCacheadoCol();
  const clave = String(cfg.admin_password || '').trim();
  if (!clave) return { ok: false, error: 'No hay clave de administrador configurada (revisa la hoja Config).' };
  if (String(password || '').trim() !== clave) {
    Utilities.sleep(400);
    return { ok: false, error: 'Clave incorrecta' };
  }
  return {
    ok: true,
    trabajadores: trabajadoresCacheadosCol(),
    menus: menusCacheadosCol(),
    pedidos: pedidosCacheadosCol(),
    platos: platosCacheadosCol(),
    config: configSinClaveCol(cfg),
    semanas: semanasCacheadasCol(),
    feriados: feriadosCacheadosCol()
  };
}

// --- TRABAJADORES ---
function guardarTrabajadorCol(data, isEdit) {
  if (!data || !data.rut || !data.nombre) return { ok: false, error: 'Falta RUT o nombre' };
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.TRABAJADORES);
    const v = h.getDataRange().getValues();
    const rn = normalizarRutCol(data.rut);
    const fila = [
      String(data.rut).trim(),
      String(data.nombre).trim(),
      data.tipo === 'Spot' ? 'Spot' : 'Fijo',
      data.activo === false || data.activo === 'No' ? 'No' : 'Si',
      data.fechaInicio || '',
      data.fechaFin || '',
      data.notas || ''
    ];
    for (let i = 1; i < v.length; i++) {
      if (normalizarRutCol(v[i][0]) === rn) {
        h.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
        cacheColInvalidar_('trabajadores');
        return { ok: true, actualizado: true };
      }
    }
    if (isEdit) return { ok: false, error: 'Trabajador no encontrado para editar' };
    h.appendRow(fila);
    cacheColInvalidar_('trabajadores');
    return { ok: true, creado: true };
  });
}
function eliminarTrabajadorCol(rut) {
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.TRABAJADORES);
    const v = h.getDataRange().getValues();
    const rn = normalizarRutCol(rut);
    for (let i = 1; i < v.length; i++) {
      if (normalizarRutCol(v[i][0]) === rn) {
        h.deleteRow(i + 1);
        cacheColInvalidar_('trabajadores');
        return { ok: true };
      }
    }
    return { ok: false, error: 'RUT no encontrado: ' + rut };
  });
}

// --- CATALOGO DE PLATOS ---
// Guarda un plato en el catalogo si no existe todavia (sin distinguir
// mayus/minus ni espacios extra), para poder elegirlo rapido despues.
// Version SIN candado propio: usarla solo desde codigo que YA sostiene el
// candado del script (ej. guardarMenuCol). Adquirir el candado dos veces en
// la misma ejecucion no esta documentado como seguro en Apps Script, asi
// que en vez de confiar en eso, guardarMenuCol llama directamente a esta
// version interna.
function guardarPlatoInterno_(nombre) {
  const limpio = String(nombre || '').trim();
  if (!limpio) return;
  const h = getHojaCol(HOJAS_COL.PLATOS);
  const v = h.getDataRange().getValues();
  const clave = limpio.toLowerCase();
  const yaExiste = v.slice(1).some(f => String(f[0]).trim().toLowerCase() === clave);
  if (!yaExiste) { h.appendRow([limpio]); cacheColInvalidar_('platos'); }
}
// Version publica con su propio candado, para cuando se llama fuera de otra
// funcion ya protegida (ej. la migracion manual agregarHojaPlatos).
function guardarPlatoCol(nombre) {
  return conCandadoCol_(() => { guardarPlatoInterno_(nombre); return null; });
}
function listarPlatosCol() {
  try {
    const v = getHojaCol(HOJAS_COL.PLATOS).getDataRange().getValues();
    return v.slice(1).map(f => String(f[0]).trim()).filter(Boolean).sort((a, b) => a.localeCompare(b));
  } catch (err) {
    return [];
  }
}
function eliminarPlatoCol(nombre) {
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.PLATOS);
    const v = h.getDataRange().getValues();
    const clave = String(nombre || '').trim().toLowerCase();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][0]).trim().toLowerCase() === clave) {
        h.deleteRow(i + 1);
        cacheColInvalidar_('platos');
        return { ok: true };
      }
    }
    return { ok: false, error: 'Plato no encontrado en el catalogo' };
  });
}

// --- MENUS ---
function guardarMenuCol(data) {
  if (!data || !data.semana || !data.dia || !data.opcion) return { ok: false, error: 'Faltan datos del menu' };
  return conCandadoCol_(() => {
    if (data.descripcion) guardarPlatoInterno_(data.descripcion);
    const h = getHojaCol(HOJAS_COL.MENUS);
    const v = h.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][0], data.semana) &&
          String(v[i][1]).trim() === String(data.dia).trim() &&
          String(v[i][2]).trim() === String(data.opcion).trim()) {
        const fila = [
          data.semana, data.dia, data.opcion,
          data.descripcion || '',
          data.activo === false ? 'No' : 'Si'
        ];
        h.getRange(i + 1, 1, 1, fila.length).setValues([fila]);
        cacheColInvalidar_('menus');
        return { ok: true, actualizado: true };
      }
    }
    // Una opcion nueva hereda el estado "especial" (y su descripcion) que ya
    // tenga ese dia, si otras opciones del mismo dia estan marcadas como
    // almuerzo mejorado.
    let especialDelDia = 'No';
    let descripcionEspecialDelDia = '';
    let semanaYaTeniaMenu = false;
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][0], data.semana)) {
        semanaYaTeniaMenu = true;
        if (String(v[i][1]).trim() === String(data.dia).trim() && String(v[i][5]).trim() === 'Si') {
          especialDelDia = 'Si';
          descripcionEspecialDelDia = String(v[i][6] || '');
        }
      }
    }
    h.appendRow([
      data.semana, data.dia, data.opcion,
      data.descripcion || '',
      data.activo === false ? 'No' : 'Si',
      especialDelDia,
      descripcionEspecialDelDia
    ]);
    cacheColInvalidar_('menus');
    // Primera opcion que se guarda para esta semana: parte en borrador, el
    // admin debe publicarla para que los trabajadores la vean. Si la
    // semana ya tenia alguna fila de Semanas (publicada o no), se respeta
    // ese estado y no se toca.
    if (!semanaYaTeniaMenu && !semanasCacheadasCol().some(f => mismaFechaCol(f.semana, data.semana))) {
      marcarSemanaPublicadaInterno_(data.semana, false);
    }
    return { ok: true, creado: true };
  });
}
function marcarDiaEspecialCol(semana, dia, especial, descripcionEspecial) {
  if (!semana || !dia) return { ok: false, error: 'Faltan datos del dia' };
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.MENUS);
    const v = h.getDataRange().getValues();
    const valor = especial ? 'Si' : 'No';
    const descripcion = especial ? String(descripcionEspecial || '') : '';
    let actualizadas = 0;
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][0], semana) && String(v[i][1]).trim() === String(dia).trim()) {
        h.getRange(i + 1, 6, 1, 2).setValues([[valor, descripcion]]);
        actualizadas++;
      }
    }
    if (!actualizadas) return { ok: false, error: 'Primero define al menos una opcion para ese dia.' };
    cacheColInvalidar_('menus');
    return { ok: true, actualizadas: actualizadas };
  });
}
function eliminarMenuCol(semana, dia, opcion) {
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.MENUS);
    const v = h.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][0], semana) &&
          String(v[i][1]).trim() === String(dia).trim() &&
          String(v[i][2]).trim() === String(opcion).trim()) {
        h.deleteRow(i + 1);
        cacheColInvalidar_('menus');
        return { ok: true };
      }
    }
    return { ok: false, error: 'Opcion de menu no encontrada' };
  });
}
// --- SEMANAS (borrador / publicada) ---
function semanasCacheadasCol() {
  return cacheColLeer_('semanas', 20, () => {
    try { return hojaAObjetosCol(HOJAS_COL.SEMANAS); } catch (err) { return []; }
  });
}
// Una semana sin fila en Semanas se considera publicada (compatibilidad
// con semanas armadas antes de este control, para no ocultarle a nadie un
// menu que ya estaba visible).
// Por defecto BORRADOR: una semana sin fila en Semanas (nunca publicada) se
// considera NO publicada. Solo se ve publicada si hay una fila explicita
// con "Si". Esto significa que cualquier semana armada antes de correr
// agregarHojaSemanas() -incluida la semana ya lista para lanzamiento- debe
// quedar marcada explicitamente "Si" en esa migracion, o el admin debe
// publicarla a mano; si no, los trabajadores no la veran.
function semanaPublicadaCol(semana) {
  const fila = semanasCacheadasCol().find(f => mismaFechaCol(f.semana, semana));
  if (!fila) return false;
  return String(fila.publicada || '').toLowerCase() === 'si';
}
// A diferencia de getHojaCol, esta crea la hoja "Semanas" sola si todavia
// no existe (ej. una planilla que venia de antes de este control y nunca
// corrio la migracion agregarHojaSemanas): asi publicar/despublicar nunca
// falla por un paso de configuracion olvidado.
function getOCrearHojaSemanasCol_() {
  const ss = ssActivaCol_();
  let h = ss.getSheetByName(HOJAS_COL.SEMANAS);
  if (!h) {
    h = ss.insertSheet(HOJAS_COL.SEMANAS);
    h.appendRow(['Semana', 'Publicada']);
  }
  return h;
}
// Version SIN candado propio, para llamar desde codigo que YA sostiene el
// candado del script (ej. guardarMenuCol al crear una semana nueva).
function marcarSemanaPublicadaInterno_(semana, publicada) {
  const h = getOCrearHojaSemanasCol_();
  const v = h.getDataRange().getValues();
  const valor = publicada ? 'Si' : 'No';
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semana)) {
      h.getRange(i + 1, 2).setValue(valor);
      cacheColInvalidar_('semanas');
      return;
    }
  }
  h.appendRow([semana, valor]);
  cacheColInvalidar_('semanas');
}
function marcarSemanaPublicadaCol(semana, publicada) {
  if (!semana) return { ok: false, error: 'Falta la semana' };
  return conCandadoCol_(() => {
    marcarSemanaPublicadaInterno_(semana, publicada);
    return { ok: true };
  });
}

// --- FERIADOS ---
// A diferencia del dia "especial", un feriado no depende de que ya exista
// una opcion de menu para ese dia: el admin puede marcarlo apenas sepa que
// no se trabaja, incluso antes de armar el menu de esa semana.
function getOCrearHojaFeriadosCol_() {
  const ss = ssActivaCol_();
  let h = ss.getSheetByName(HOJAS_COL.FERIADOS);
  if (!h) {
    h = ss.insertSheet(HOJAS_COL.FERIADOS);
    h.appendRow(['Semana', 'Dia']);
  }
  return h;
}
function diaEsFeriadoCol(semana, dia) {
  return feriadosCacheadosCol().some(f =>
    mismaFechaCol(f.semana, semana) && String(f.dia).trim() === String(dia).trim());
}
function marcarFeriadoCol(semana, dia, feriado) {
  if (!semana || !dia) return { ok: false, error: 'Faltan datos del dia' };
  return conCandadoCol_(() => {
    const h = getOCrearHojaFeriadosCol_();
    const v = h.getDataRange().getValues();
    let yaEstaba = false;
    for (let i = v.length - 1; i >= 1; i--) {
      if (mismaFechaCol(v[i][0], semana) && String(v[i][1]).trim() === String(dia).trim()) {
        yaEstaba = true;
        if (!feriado) h.deleteRow(i + 1);
      }
    }
    if (feriado && !yaEstaba) h.appendRow([semana, dia]);
    cacheColInvalidar_('feriados');
    let pedidosBorrados = 0;
    // Un feriado significa que ese dia no hay colacion: cualquier pedido que
    // ya existiera ahi (de antes de marcarlo feriado) queda sin sentido y se
    // borra, para que el conteo/reporte no arrastre pedidos de un dia que no
    // se va a trabajar.
    if (feriado) {
      const hp = getHojaCol(HOJAS_COL.PEDIDOS);
      const vp = hp.getDataRange().getValues();
      for (let i = vp.length - 1; i >= 1; i--) {
        if (mismaFechaCol(vp[i][1], semana) && String(vp[i][4]).trim() === String(dia).trim()) {
          hp.deleteRow(i + 1);
          pedidosBorrados++;
        }
      }
      if (pedidosBorrados) cacheColInvalidar_('pedidos');
    }
    return { ok: true, pedidosBorrados: pedidosBorrados };
  });
}

// --- PEDIDOS ---
// Un trabajador solo puede tener UNA opcion elegida por dia dentro de la
// misma semana: si ya existe una fila para (semana, rut, dia) se actualiza
// en vez de duplicar. Esto es lo que impide las anotaciones dobles.
function guardarPedidoCol(data) {
  if (!data || !data.semana || !data.rut || !data.dia || !data.opcion) {
    return { ok: false, error: 'Faltan datos del pedido' };
  }
  if (!data.asAdmin && inscripcionesCerradasCol(data.semana)) {
    return { ok: false, error: 'El plazo para anotarse a esta semana ya cerro. Si necesitas hacer un cambio, contacta a administracion.', cerrado: true };
  }
  if (!data.asAdmin && !semanaPublicadaCol(data.semana)) {
    return { ok: false, error: 'El menu de esta semana todavia no esta publicado.' };
  }
  if (!data.asAdmin && diaEsFeriadoCol(data.semana, data.dia)) {
    return { ok: false, error: 'Ese día es feriado, no hay colación.' };
  }
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.PEDIDOS);
    const v = h.getDataRange().getValues();
    const rn = normalizarRutCol(data.rut);
    const ahora = timestampLocalCol();
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][1], data.semana) &&
          normalizarRutCol(v[i][2]) === rn &&
          String(v[i][4]).trim() === String(data.dia).trim()) {
        h.getRange(i + 1, 6, 1, 2).setValues([[data.opcion, ahora]]);
        cacheColInvalidar_('pedidos');
        return { ok: true, actualizado: true };
      }
    }
    const id = Utilities.getUuid();
    h.appendRow([id, data.semana, data.rut, data.nombre || '', data.dia, data.opcion, ahora]);
    cacheColInvalidar_('pedidos');
    return { ok: true, creado: true, id: id };
  });
}
function eliminarPedidoCol(semana, rut, dia, asAdmin) {
  if (!asAdmin && inscripcionesCerradasCol(semana)) {
    return { ok: false, error: 'El plazo para anotarse a esta semana ya cerro. Si necesitas hacer un cambio, contacta a administracion.', cerrado: true };
  }
  return conCandadoCol_(() => {
    const h = getHojaCol(HOJAS_COL.PEDIDOS);
    const v = h.getDataRange().getValues();
    const rn = normalizarRutCol(rut);
    for (let i = 1; i < v.length; i++) {
      if (mismaFechaCol(v[i][1], semana) &&
          normalizarRutCol(v[i][2]) === rn &&
          String(v[i][4]).trim() === String(dia).trim()) {
        h.deleteRow(i + 1);
        cacheColInvalidar_('pedidos');
        return { ok: true };
      }
    }
    return { ok: false, error: 'Pedido no encontrado' };
  });
}

// --- TRABAJADOR: login + su semana ---
function loginTrabajadorCol(rutIngresado) {
  if (!rutIngresado) return { ok: false, error: 'Ingresa tu RUT' };
  const rn = normalizarRutCol(rutIngresado);
  const trabajadores = trabajadoresCacheadosCol();
  const t = trabajadores.find(x => normalizarRutCol(x.rut) === rn);
  if (!t) { Utilities.sleep(300); return { ok: false, error: 'RUT no encontrado. Consulta con administracion.' }; }
  if (String(t.activo).toLowerCase() !== 'si') {
    return { ok: false, error: 'Tu registro esta inactivo. Consulta con administracion.' };
  }
  const misPedidos = pedidosCacheadosCol().filter(p => normalizarRutCol(p.rut) === rn);
  return {
    ok: true,
    trabajador: { rut: String(t.rut), nombre: String(t.nombre), tipo: String(t.tipo || 'Fijo') },
    menus: menusPublicadosCol(menusCacheadosCol()),
    pedidos: misPedidos,
    config: configSinClaveCol(configCacheadoCol()),
    feriados: feriadosCacheadosCol()
  };
}

// --- UTILIDADES DE DIAGNOSTICO (ejecutar manualmente desde el editor) ---
function listarHojasCol() {
  const nombres = ssActivaCol_().getSheets().map(s => s.getName());
  Logger.log('Hojas encontradas: ' + JSON.stringify(nombres));
  return nombres;
}
function testAPICol() {
  const d = obtenerTodoCol();
  Logger.log('Trabajadores:' + d.trabajadores.length + ' Menus:' + d.menus.length + ' Pedidos:' + d.pedidos.length);
}
function crearHojasIniciales() {
  const ss = ssActivaCol_();
  const specs = {
    Trabajadores: ['RUT', 'Nombre', 'Tipo', 'Activo', 'FechaInicio', 'FechaFin', 'Notas'],
    Menus: ['Semana', 'Dia', 'Opcion', 'Descripcion', 'Activo', 'Especial', 'DescripcionEspecial'],
    Pedidos: ['ID', 'Semana', 'RUT', 'Nombre', 'Dia', 'Opcion', 'Timestamp'],
    Config: ['Clave', 'Valor'],
    Platos: ['Nombre'],
    Semanas: ['Semana', 'Publicada'],
    Feriados: ['Semana', 'Dia']
  };
  Object.keys(specs).forEach(nombre => {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) h.appendRow(specs[nombre]);
  });
  const cfg = ss.getSheetByName('Config');
  const v = cfg.getDataRange().getValues();
  const claves = v.slice(1).map(f => String(f[0]).toLowerCase().trim());
  if (!claves.includes('admin_password')) cfg.appendRow(['admin_password', 'cambiar123']);
  if (!claves.includes('cierre_dias_antes')) cfg.appendRow(['cierre_dias_antes', 5]);
  if (!claves.includes('cierre_hora')) actualizarConfigCol('cierre_hora', '14:00');
  Logger.log('Hojas listas. Recuerda cambiar la clave admin_password en la hoja Config.');
}
// Ejecutar UNA VEZ si tu hoja "Menus" ya existia antes de que se agregara la
// columna "Especial" (almuerzo mejorado). Agrega el encabezado si falta.
function agregarColumnaEspecial() {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const encabezado = h.getRange(1, 1, 1, Math.max(6, h.getLastColumn())).getValues()[0];
  if (String(encabezado[5] || '').trim().toLowerCase() === 'especial') {
    Logger.log('La columna Especial ya existe.');
    return;
  }
  h.getRange(1, 6).setValue('Especial');
  Logger.log('Columna "Especial" agregada en Menus!F1.');
}
// Ejecutar UNA VEZ si tu hoja "Menus" ya existia antes de que se agregara
// la descripcion del almuerzo mejorado (ej. "Torta con bebida").
function agregarColumnaDescripcionEspecial() {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const encabezado = h.getRange(1, 1, 1, Math.max(7, h.getLastColumn())).getValues()[0];
  if (String(encabezado[6] || '').trim().toLowerCase() === 'descripcionespecial') {
    Logger.log('La columna DescripcionEspecial ya existe.');
    return;
  }
  h.getRange(1, 7).setValue('DescripcionEspecial');
  Logger.log('Columna "DescripcionEspecial" agregada en Menus!G1.');
}
// Ejecutar UNA VEZ si tu planilla ya existia antes de que se agregara el
// cierre de inscripciones configurable. Agrega las claves con sus valores
// por defecto (5 dias antes del lunes, a las 14:00) si no existen.
function agregarConfigCierre() {
  const h = getHojaCol(HOJAS_COL.CONFIG);
  const v = h.getDataRange().getValues();
  const claves = v.slice(1).map(f => String(f[0]).toLowerCase().trim());
  if (!claves.includes('cierre_dias_antes')) h.appendRow(['cierre_dias_antes', 5]);
  if (!claves.includes('cierre_hora')) actualizarConfigCol('cierre_hora', '14:00');
  Logger.log('Config de cierre lista (cierre_dias_antes, cierre_hora).');
}
// Ejecutar UNA VEZ si tu celda "cierre_hora" en Config ya quedo corrompida
// (Sheets la convirtio en fecha/hora en vez de dejarla como texto "14:00").
// La relee, rescata la hora:minuto que tenga guardada y la vuelve a
// escribir forzando formato de texto para que no se corrompa de nuevo.
function repararCierreHora() {
  const cfg = obtenerConfigCol(); // ya normaliza Date -> "HH:mm"
  const horaActual = cfg.cierre_hora || '14:00';
  actualizarConfigCol('cierre_hora', horaActual);
  Logger.log('cierre_hora reparada como texto: ' + horaActual);
}
// Ejecutar UNA VEZ si tu planilla ya existia antes del catalogo de platos.
// Crea la hoja "Platos" si falta y la precarga con todas las descripciones
// ya usadas en "Menus", para no perder lo que ya tenias escrito.
function agregarHojaPlatos() {
  const ss = ssActivaCol_();
  let h = ss.getSheetByName(HOJAS_COL.PLATOS);
  if (!h) h = ss.insertSheet(HOJAS_COL.PLATOS);
  if (h.getLastRow() === 0) h.appendRow(['Nombre']);
  const menus = menusAObjetosCol();
  let agregados = 0;
  menus.forEach(m => {
    if (m.descripcion) {
      const antes = listarPlatosCol().length;
      guardarPlatoCol(m.descripcion);
      if (listarPlatosCol().length > antes) agregados++;
    }
  });
  Logger.log('Hoja Platos lista. Platos agregados desde Menus: ' + agregados);
}
// Ejecutar UNA VEZ si tu planilla ya existia antes del control de
// publicacion por semana. Crea la hoja "Semanas" si falta y marca como
// publicadas (Si) todas las semanas que ya tengan menu armado, para que
// nadie pierda de vista un menu que los trabajadores ya podian ver.
function agregarHojaSemanas() {
  const ss = ssActivaCol_();
  let h = ss.getSheetByName(HOJAS_COL.SEMANAS);
  if (!h) h = ss.insertSheet(HOJAS_COL.SEMANAS);
  if (h.getLastRow() === 0) h.appendRow(['Semana', 'Publicada']);
  const semanasExistentes = new Set(
    hojaAObjetosCol(HOJAS_COL.SEMANAS).map(f => formatearFechaCol(f.semana))
  );
  const semanasConMenu = new Set(menusAObjetosCol().map(m => formatearFechaCol(m.semana)));
  let agregadas = 0;
  semanasConMenu.forEach(semana => {
    if (!semanasExistentes.has(semana)) {
      h.appendRow([semana, 'Si']);
      agregadas++;
    }
  });
  cacheColInvalidar_('semanas');
  Logger.log('Hoja Semanas lista. Semanas marcadas como publicadas: ' + agregadas);
}
