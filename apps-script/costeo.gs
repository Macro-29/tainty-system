/* =========================================================
   Tainty Labs — Apps Script: COSTEO POR PRODUCTO
   Backend para costeo_tainty.html
   ---------------------------------------------------------
   Cambios respecto a la versión anterior:
   A) Batch writes con getRange().setValues() (1 escritura por save,
      en lugar de N appendRow secuenciales). 5-10× más rápido.
   B) LockService para evitar que dos pestañas escriban a la vez.
   C) Match case-insensitive al borrar filas previas del producto.
   D) Más metadata en la respuesta (filas insertadas, duración, version)
      para que el cliente verifique que se escribió todo.
   E) Logger en cada acción + un endpoint `?action=ping` para verificar
      que el deployment está activo y la hoja accesible.
   F) Validaciones defensivas (data vacío, hoja inexistente, insumos sin
      nombre, etc.) que ahora devuelven mensajes claros en vez de fallar.
   ========================================================= */

const COSTEO_ID = '1U9xDfX1wkiSiL8KsJ_Qb9lwGUtmX97VBLfnLwUCEWJs';
const COSTEO_SHEET = 'COSTO POR PRODUCTO';
const SCRIPT_VERSION = '2026-05-13';

function getTipoConfig(tipo) {
  if (['INSUMO', 'LAINA'].indexOf(tipo) >= 0) return 'kg';
  if (['CAPSULA', 'ENVASE'].indexOf(tipo) >= 0) return 'und';
  return 'fijo'; // ETIQUETA, CAJA
}

/* =========================================================
   Entry points: doGet (JSONP) y doPost (FormData no-cors)
   ========================================================= */

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  try {
    if (e && e.parameter && e.parameter.data) {
      const data = JSON.parse(decodeURIComponent(e.parameter.data));
      return dispatchAction(e, data);
    }
    // Sin `data` → leer catalogo. Soporta también ?action=ping para healthcheck.
    if (e && e.parameter && e.parameter.action === 'ping') return ping(e);
    return leerCatalogo(e);
  } catch (err) {
    console.error('doGet error:', err.message, err.stack);
    return respuestaJsonp(e, {
      success: false,
      error: 'doGet: ' + err.message,
      stack: err.stack
    });
  }
}

function doPost(e) {
  try {
    let data;
    if (e.parameter && e.parameter.data) data = JSON.parse(e.parameter.data);
    else if (e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    if (!data) return buildCosteoResponse({ success: false, error: 'No data received' });
    return dispatchAction(e, data);
  } catch (err) {
    console.error('doPost error:', err.message, err.stack);
    return buildCosteoResponse({ success: false, error: err.message, stack: err.stack });
  }
}

function dispatchAction(e, data) {
  if (!data || !data.action) {
    return respuestaJsonp(e, { success: false, error: 'Falta el campo "action"' });
  }
  console.log('dispatchAction:', data.action, '· producto:', data.producto);
  switch (data.action) {
    case 'guardarProducto':  return guardarProducto(e, data);
    case 'eliminarProducto': return eliminarProducto(e, data);
    case 'ping':             return ping(e);
    default:
      return respuestaJsonp(e, {
        success: false,
        error: 'Accion no reconocida: ' + data.action
      });
  }
}

/* =========================================================
   Acciones
   ========================================================= */

/**
 * Guarda un producto y sus insumos. Borra TODAS las filas previas con el
 * mismo nombre (case-insensitive) e inserta las nuevas en una sola escritura.
 * Usa LockService para evitar concurrencia.
 */
function guardarProducto(e, data) {
  const t0 = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return respuestaJsonp(e, {
      success: false,
      error: 'Hoja ocupada (otra escritura en curso). Reintenta en un momento.'
    });
  }
  try {
    if (!data.producto || !data.producto.toString().trim()) {
      return respuestaJsonp(e, { success: false, error: 'Falta el nombre del producto' });
    }
    if (!Array.isArray(data.insumos) || data.insumos.length === 0) {
      return respuestaJsonp(e, { success: false, error: 'El producto no tiene insumos' });
    }

    const sheet = getCosteoSheet_();
    if (!sheet) {
      return respuestaJsonp(e, {
        success: false,
        error: 'Hoja no encontrada: "' + COSTEO_SHEET + '" en spreadsheet ' + COSTEO_ID
      });
    }

    const producto = data.producto.toString().trim();
    const productoUp = producto.toUpperCase();

    // (C) Borrar filas previas — case-insensitive
    const allData = sheet.getDataRange().getValues();
    let filasBorradas = 0;
    for (let i = allData.length - 1; i >= 1; i--) {
      if ((allData[i][0] || '').toString().trim().toUpperCase() === productoUp) {
        sheet.deleteRow(i + 1);
        filasBorradas++;
      }
    }

    // Construir filas en memoria (no escribir todavía).
    const filas = data.insumos
      .filter(ins => ins && ins.nombre && ins.nombre.toString().trim())
      .map(ins => buildFila_(producto, ins));

    if (filas.length === 0) {
      SpreadsheetApp.flush(); // confirma los deletes
      return respuestaJsonp(e, {
        success: false,
        error: 'Ningún insumo tenía nombre válido',
        filasBorradas: filasBorradas
      });
    }

    // (A) Batch write — una sola llamada en vez de N appendRow
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, filas.length, 13).setValues(filas);
    SpreadsheetApp.flush();

    const duracion = Date.now() - t0;
    console.log('guardarProducto OK:', producto, '·', filas.length, 'filas ·', duracion, 'ms');

    return respuestaJsonp(e, {
      success: true,
      producto: producto,
      filasInsertadas: filas.length,
      filasBorradas: filasBorradas,
      duracionMs: duracion,
      version: SCRIPT_VERSION
    });
  } catch (err) {
    console.error('guardarProducto error:', err.message, err.stack);
    return respuestaJsonp(e, {
      success: false,
      error: 'guardarProducto: ' + err.message,
      stack: err.stack,
      producto: data.producto
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Construye la fila [producto, tipo, ...] para un insumo según su config.
 * Auto-calcula costos = cantidades × precioUnitario para tipos kg/und.
 */
function buildFila_(producto, ins) {
  const tipo = (ins.tipo || '').toString().toUpperCase();
  const config = getTipoConfig(tipo);

  let precio = 0;
  const cant = { 100: 0, 300: 0, 500: 0, 1000: 0 };
  const costos = { 100: 0, 300: 0, 500: 0, 1000: 0 };

  if (config === 'kg' || config === 'und') {
    precio = parseFloat(ins.precioUnitario) || parseFloat(ins.costoPorKg) || parseFloat(ins.costoUnd) || 0;
    if (ins.cantidades) {
      cant[100]  = parseFloat(ins.cantidades[100])  || 0;
      cant[300]  = parseFloat(ins.cantidades[300])  || 0;
      cant[500]  = parseFloat(ins.cantidades[500])  || 0;
      cant[1000] = parseFloat(ins.cantidades[1000]) || 0;
    }
    // Costos derivados
    costos[100]  = cant[100]  * precio;
    costos[300]  = cant[300]  * precio;
    costos[500]  = cant[500]  * precio;
    costos[1000] = cant[1000] * precio;
  } else if (ins.costos) {
    costos[100]  = parseFloat(ins.costos[100])  || 0;
    costos[300]  = parseFloat(ins.costos[300])  || 0;
    costos[500]  = parseFloat(ins.costos[500])  || 0;
    costos[1000] = parseFloat(ins.costos[1000]) || 0;
  }

  return [
    producto, tipo, (ins.nombre || '').toString().trim(), (ins.proveedor || '').toString().trim(),
    precio,
    cant[100], cant[300], cant[500], cant[1000],
    costos[100], costos[300], costos[500], costos[1000]
  ];
}

function eliminarProducto(e, data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return respuestaJsonp(e, { success: false, error: 'Hoja ocupada, reintenta' });
  }
  try {
    if (!data.producto) {
      return respuestaJsonp(e, { success: false, error: 'Falta el nombre del producto' });
    }
    const sheet = getCosteoSheet_();
    if (!sheet) return respuestaJsonp(e, { success: false, error: 'Hoja no encontrada' });

    const allData = sheet.getDataRange().getValues();
    const objetivo = data.producto.toString().trim().toUpperCase();
    let filasBorradas = 0;
    for (let i = allData.length - 1; i >= 1; i--) {
      if ((allData[i][0] || '').toString().trim().toUpperCase() === objetivo) {
        sheet.deleteRow(i + 1);
        filasBorradas++;
      }
    }
    SpreadsheetApp.flush();
    console.log('eliminarProducto OK:', data.producto, '·', filasBorradas, 'filas');
    return respuestaJsonp(e, {
      success: true,
      eliminado: data.producto,
      filasBorradas: filasBorradas
    });
  } catch (err) {
    console.error('eliminarProducto error:', err.message, err.stack);
    return respuestaJsonp(e, {
      success: false,
      error: 'eliminarProducto: ' + err.message,
      stack: err.stack
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Devuelve el catálogo completo agrupado por producto. Estructura:
 * [ { nombre, insumos: [ { tipo, nombre, proveedor, precioUnitario?, cantidades?, costos? } ] }, ... ]
 */
function leerCatalogo(e) {
  const sheet = getCosteoSheet_();
  if (!sheet) {
    return respuestaJsonp(e, {
      success: false,
      error: 'Hoja no encontrada: "' + COSTEO_SHEET + '"'
    });
  }
  const data = sheet.getDataRange().getValues();
  const catalogo = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const producto = (r[0] || '').toString().trim();
    const tipo = (r[1] || '').toString().trim().toUpperCase();
    const nombre = (r[2] || '').toString().trim();
    const proveedor = (r[3] || '').toString().trim();
    const precio = parseFloat(r[4]) || 0;
    const c100 = parseFloat(r[5])  || 0;
    const c300 = parseFloat(r[6])  || 0;
    const c500 = parseFloat(r[7])  || 0;
    const c1000 = parseFloat(r[8]) || 0;
    const k100 = parseFloat(r[9])  || 0;
    const k300 = parseFloat(r[10]) || 0;
    const k500 = parseFloat(r[11]) || 0;
    const k1000 = parseFloat(r[12])|| 0;
    if (!producto || !tipo || !nombre) continue;
    if (!catalogo[producto]) catalogo[producto] = { nombre: producto, insumos: [] };

    const config = getTipoConfig(tipo);
    const insumo = { tipo: tipo, nombre: nombre, proveedor: proveedor };
    if (config === 'kg' || config === 'und') {
      insumo.precioUnitario = precio;
      insumo.cantidades = { 100: c100, 300: c300, 500: c500, 1000: c1000 };
    } else {
      // Para ETIQUETA/CAJA usar la columna de costos. Fallback a cantidades por compatibilidad.
      if (k100 || k300 || k500 || k1000) {
        insumo.costos = { 100: k100, 300: k300, 500: k500, 1000: k1000 };
      } else {
        insumo.costos = { 100: c100, 300: c300, 500: c500, 1000: c1000 };
      }
    }
    catalogo[producto].insumos.push(insumo);
  }
  return respuestaJsonp(e, {
    success: true,
    data: Object.values(catalogo),
    version: SCRIPT_VERSION
  });
}

/**
 * Healthcheck. Verifica que el deployment responde y la hoja es accesible.
 * Útil para depurar desde la consola del navegador:
 *   fetch(COSTEO_URL + '?action=ping&callback=cb').then(r=>r.text()).then(console.log)
 */
function ping(e) {
  const t0 = Date.now();
  let sheetOk = false, filas = 0, sheetError = null;
  try {
    const sheet = getCosteoSheet_();
    if (sheet) {
      sheetOk = true;
      filas = Math.max(0, sheet.getLastRow() - 1);
    } else {
      sheetError = 'Hoja "' + COSTEO_SHEET + '" no encontrada';
    }
  } catch (err) {
    sheetError = err.message;
  }
  return respuestaJsonp(e, {
    success: true,
    pong: true,
    version: SCRIPT_VERSION,
    spreadsheetId: COSTEO_ID,
    sheetName: COSTEO_SHEET,
    sheetOk: sheetOk,
    sheetError: sheetError,
    filasEnHoja: filas,
    duracionMs: Date.now() - t0
  });
}

/* =========================================================
   Helpers internos
   ========================================================= */

function getCosteoSheet_() {
  const ss = SpreadsheetApp.openById(COSTEO_ID);
  return ss.getSheetByName(COSTEO_SHEET); // null si no existe
}

function respuestaJsonp(e, obj) {
  const result = JSON.stringify(obj);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + result + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}

function buildCosteoResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================
   Utilidades de mantenimiento (ejecutar manualmente desde
   el editor de Apps Script, NO se exponen como endpoints)
   ========================================================= */

/** Inicializa o repara la fila de cabeceras de la hoja. */
function inicializarColumnasNuevas() {
  const sheet = getCosteoSheet_();
  if (!sheet) throw new Error('Hoja no encontrada: ' + COSTEO_SHEET);
  const cabeceras = sheet.getRange(1, 1, 1, 13).getValues()[0];
  const nuevas = [
    cabeceras[0]  || 'Producto',
    cabeceras[1]  || 'Tipo',
    cabeceras[2]  || 'Descripción',
    cabeceras[3]  || 'Proveedor',
    cabeceras[4]  || 'Precio unitario',
    cabeceras[5]  || 'Cant 100',
    cabeceras[6]  || 'Cant 300',
    cabeceras[7]  || 'Cant 500',
    cabeceras[8]  || 'Cant 1000',
    cabeceras[9]  || 'Costo 100',
    cabeceras[10] || 'Costo 300',
    cabeceras[11] || 'Costo 500',
    cabeceras[12] || 'Costo 1000'
  ];
  sheet.getRange(1, 1, 1, 13).setValues([nuevas]);
  console.log('Columnas inicializadas');
}

/**
 * Rellena las columnas J-M (Costo 100/300/500/1000) de los productos
 * existentes según: cantidad × precio unitario (para kg/und).
 * No toca ETIQUETA/CAJA (costos ya están ingresados manualmente).
 */
function recalcularCostosExistentes() {
  const sheet = getCosteoSheet_();
  if (!sheet) throw new Error('Hoja no encontrada: ' + COSTEO_SHEET);
  const data = sheet.getDataRange().getValues();
  const filasActualizar = [];
  const indicesActualizar = [];

  for (let i = 1; i < data.length; i++) {
    const tipo = (data[i][1] || '').toString().trim().toUpperCase();
    if (!tipo) continue;
    if (getTipoConfig(tipo) === 'fijo') continue;
    const precio = parseFloat(data[i][4]) || 0;
    const c100  = parseFloat(data[i][5]) || 0;
    const c300  = parseFloat(data[i][6]) || 0;
    const c500  = parseFloat(data[i][7]) || 0;
    const c1000 = parseFloat(data[i][8]) || 0;
    filasActualizar.push([c100 * precio, c300 * precio, c500 * precio, c1000 * precio]);
    indicesActualizar.push(i + 1);
  }

  // Escribir cada bloque contiguo de filas en una sola operación batch.
  // (Optimización: si las filas a actualizar son consecutivas, una sola setValues.)
  let actualizadas = 0;
  let bloque = [];
  let bloqueInicio = null;
  for (let k = 0; k < indicesActualizar.length; k++) {
    const idx = indicesActualizar[k];
    if (bloqueInicio === null) { bloqueInicio = idx; bloque = [filasActualizar[k]]; }
    else if (idx === bloqueInicio + bloque.length) { bloque.push(filasActualizar[k]); }
    else {
      sheet.getRange(bloqueInicio, 10, bloque.length, 4).setValues(bloque);
      actualizadas += bloque.length;
      bloqueInicio = idx;
      bloque = [filasActualizar[k]];
    }
  }
  if (bloque.length > 0) {
    sheet.getRange(bloqueInicio, 10, bloque.length, 4).setValues(bloque);
    actualizadas += bloque.length;
  }

  SpreadsheetApp.flush();
  console.log('Costos recalculados en ' + actualizadas + ' filas');
  return { success: true, actualizadas: actualizadas };
}
