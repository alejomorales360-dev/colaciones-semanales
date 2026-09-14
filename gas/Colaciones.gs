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

const HOJAS_COL = {
  TRABAJADORES: 'Trabajadores',
  MENUS: 'Menus',
  PEDIDOS: 'Pedidos',
  CONFIG: 'Config',
  PLATOS: 'Platos'
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

function procesarAccionCol(body) {
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
    case 'copiarMenuSemana':
      return copiarMenuSemanaCol(body.semanaOrigen, body.semanaDestino);
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

// --- HELPERS ---
function getHojaCol(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
function obtenerTodoCol() {
  return {
    ok: true,
    trabajadores: hojaAObjetosCol(HOJAS_COL.TRABAJADORES),
    menus: hojaAObjetosCol(HOJAS_COL.MENUS),
    pedidos: hojaAObjetosCol(HOJAS_COL.PEDIDOS),
    platos: listarPlatosCol(),
    config: obtenerConfigCol(),
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
  const h = getHojaCol(HOJAS_COL.CONFIG);
  const v = h.getDataRange().getValues();
  // Forzar formato de texto en la celda para que Sheets no reinterprete
  // valores como "14:00" como una hora/fecha (lo que corrompia cierre_hora).
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).toLowerCase().trim() === String(clave).toLowerCase().trim()) {
      h.getRange(i + 1, 2).setNumberFormat('@').setValue(valor);
      return { ok: true };
    }
  }
  const fila = h.getLastRow() + 1;
  h.getRange(fila, 1).setValue(clave);
  h.getRange(fila, 2).setNumberFormat('@').setValue(valor);
  return { ok: true, creado: true };
}
// Calcula el instante (Date) en que se cierran las inscripciones de una
// semana: "cierre_dias_antes" dias antes del lunes de esa semana, a la hora
// "cierre_hora". Por defecto: miercoles de la semana previa a las 14:00.
function calcularCierreCol(semana, cfg) {
  cfg = cfg || obtenerConfigCol();
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
  const cfg = obtenerConfigCol();
  const clave = String(cfg.admin_password || '').trim();
  if (!clave) return { ok: false, error: 'No hay clave de administrador configurada (revisa la hoja Config).' };
  if (String(password || '').trim() !== clave) {
    Utilities.sleep(400);
    return { ok: false, error: 'Clave incorrecta' };
  }
  return {
    ok: true,
    trabajadores: hojaAObjetosCol(HOJAS_COL.TRABAJADORES),
    menus: hojaAObjetosCol(HOJAS_COL.MENUS),
    pedidos: hojaAObjetosCol(HOJAS_COL.PEDIDOS),
    platos: listarPlatosCol(),
    config: cfg
  };
}

// --- TRABAJADORES ---
function guardarTrabajadorCol(data, isEdit) {
  if (!data || !data.rut || !data.nombre) return { ok: false, error: 'Falta RUT o nombre' };
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
      return { ok: true, actualizado: true };
    }
  }
  if (isEdit) return { ok: false, error: 'Trabajador no encontrado para editar' };
  h.appendRow(fila);
  return { ok: true, creado: true };
}
function eliminarTrabajadorCol(rut) {
  const h = getHojaCol(HOJAS_COL.TRABAJADORES);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(rut);
  for (let i = 1; i < v.length; i++) {
    if (normalizarRutCol(v[i][0]) === rn) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'RUT no encontrado: ' + rut };
}

// --- CATALOGO DE PLATOS ---
// Guarda un plato en el catalogo si no existe todavia (sin distinguir
// mayus/minus ni espacios extra), para poder elegirlo rapido despues.
function guardarPlatoCol(nombre) {
  const limpio = String(nombre || '').trim();
  if (!limpio) return;
  const h = getHojaCol(HOJAS_COL.PLATOS);
  const v = h.getDataRange().getValues();
  const clave = limpio.toLowerCase();
  const yaExiste = v.slice(1).some(f => String(f[0]).trim().toLowerCase() === clave);
  if (!yaExiste) h.appendRow([limpio]);
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
  const h = getHojaCol(HOJAS_COL.PLATOS);
  const v = h.getDataRange().getValues();
  const clave = String(nombre || '').trim().toLowerCase();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === clave) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Plato no encontrado en el catalogo' };
}

// --- MENUS ---
function guardarMenuCol(data) {
  if (!data || !data.semana || !data.dia || !data.opcion) return { ok: false, error: 'Faltan datos del menu' };
  if (data.descripcion) guardarPlatoCol(data.descripcion);
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
      return { ok: true, actualizado: true };
    }
  }
  // Una opcion nueva hereda el estado "especial" (y su descripcion) que ya
  // tenga ese dia, si otras opciones del mismo dia estan marcadas como
  // almuerzo mejorado.
  let especialDelDia = 'No';
  let descripcionEspecialDelDia = '';
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], data.semana) && String(v[i][1]).trim() === String(data.dia).trim() && String(v[i][5]).trim() === 'Si') {
      especialDelDia = 'Si';
      descripcionEspecialDelDia = String(v[i][6] || '');
      break;
    }
  }
  h.appendRow([
    data.semana, data.dia, data.opcion,
    data.descripcion || '',
    data.activo === false ? 'No' : 'Si',
    especialDelDia,
    descripcionEspecialDelDia
  ]);
  return { ok: true, creado: true };
}
function marcarDiaEspecialCol(semana, dia, especial, descripcionEspecial) {
  if (!semana || !dia) return { ok: false, error: 'Faltan datos del dia' };
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
  return { ok: true, actualizadas: actualizadas };
}
function eliminarMenuCol(semana, dia, opcion) {
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semana) &&
        String(v[i][1]).trim() === String(dia).trim() &&
        String(v[i][2]).trim() === String(opcion).trim()) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Opcion de menu no encontrada' };
}
function copiarMenuSemanaCol(semanaOrigen, semanaDestino) {
  if (!semanaOrigen || !semanaDestino) return { ok: false, error: 'Faltan semanas' };
  const h = getHojaCol(HOJAS_COL.MENUS);
  const v = h.getDataRange().getValues();
  let copiadas = 0;
  const nuevasFilas = [];
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][0], semanaOrigen)) {
      nuevasFilas.push([semanaDestino, v[i][1], v[i][2], v[i][3], v[i][4], v[i][5] || 'No']);
      copiadas++;
    }
  }
  nuevasFilas.forEach(f => h.appendRow(f));
  return { ok: true, copiadas: copiadas };
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
  const h = getHojaCol(HOJAS_COL.PEDIDOS);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(data.rut);
  const ahora = timestampLocalCol();
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][1], data.semana) &&
        normalizarRutCol(v[i][2]) === rn &&
        String(v[i][4]).trim() === String(data.dia).trim()) {
      h.getRange(i + 1, 6, 1, 2).setValues([[data.opcion, ahora]]);
      return { ok: true, actualizado: true };
    }
  }
  const id = Utilities.getUuid();
  h.appendRow([id, data.semana, data.rut, data.nombre || '', data.dia, data.opcion, ahora]);
  return { ok: true, creado: true, id: id };
}
function eliminarPedidoCol(semana, rut, dia, asAdmin) {
  if (!asAdmin && inscripcionesCerradasCol(semana)) {
    return { ok: false, error: 'El plazo para anotarse a esta semana ya cerro. Si necesitas hacer un cambio, contacta a administracion.', cerrado: true };
  }
  const h = getHojaCol(HOJAS_COL.PEDIDOS);
  const v = h.getDataRange().getValues();
  const rn = normalizarRutCol(rut);
  for (let i = 1; i < v.length; i++) {
    if (mismaFechaCol(v[i][1], semana) &&
        normalizarRutCol(v[i][2]) === rn &&
        String(v[i][4]).trim() === String(dia).trim()) {
      h.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Pedido no encontrado' };
}

// --- TRABAJADOR: login + su semana ---
function loginTrabajadorCol(rutIngresado) {
  if (!rutIngresado) return { ok: false, error: 'Ingresa tu RUT' };
  const rn = normalizarRutCol(rutIngresado);
  const trabajadores = hojaAObjetosCol(HOJAS_COL.TRABAJADORES);
  const t = trabajadores.find(x => normalizarRutCol(x.rut) === rn);
  if (!t) { Utilities.sleep(300); return { ok: false, error: 'RUT no encontrado. Consulta con administracion.' }; }
  if (String(t.activo).toLowerCase() !== 'si') {
    return { ok: false, error: 'Tu registro esta inactivo. Consulta con administracion.' };
  }
  const misPedidos = hojaAObjetosCol(HOJAS_COL.PEDIDOS).filter(p => normalizarRutCol(p.rut) === rn);
  return {
    ok: true,
    trabajador: { rut: String(t.rut), nombre: String(t.nombre), tipo: String(t.tipo || 'Fijo') },
    menus: hojaAObjetosCol(HOJAS_COL.MENUS),
    pedidos: misPedidos,
    config: obtenerConfigCol()
  };
}

// --- UTILIDADES DE DIAGNOSTICO (ejecutar manualmente desde el editor) ---
function listarHojasCol() {
  const nombres = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
  Logger.log('Hojas encontradas: ' + JSON.stringify(nombres));
  return nombres;
}
function testAPICol() {
  const d = obtenerTodoCol();
  Logger.log('Trabajadores:' + d.trabajadores.length + ' Menus:' + d.menus.length + ' Pedidos:' + d.pedidos.length);
}
function crearHojasIniciales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const specs = {
    Trabajadores: ['RUT', 'Nombre', 'Tipo', 'Activo', 'FechaInicio', 'FechaFin', 'Notas'],
    Menus: ['Semana', 'Dia', 'Opcion', 'Descripcion', 'Activo', 'Especial', 'DescripcionEspecial'],
    Pedidos: ['ID', 'Semana', 'RUT', 'Nombre', 'Dia', 'Opcion', 'Timestamp'],
    Config: ['Clave', 'Valor'],
    Platos: ['Nombre']
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName(HOJAS_COL.PLATOS);
  if (!h) h = ss.insertSheet(HOJAS_COL.PLATOS);
  if (h.getLastRow() === 0) h.appendRow(['Nombre']);
  const menus = hojaAObjetosCol(HOJAS_COL.MENUS);
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
