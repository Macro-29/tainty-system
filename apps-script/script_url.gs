/* =========================================================
   Tainty Labs — Apps Script principal (SCRIPT_URL)
   ---------------------------------------------------------
   Backend de: formulario_tainty, pagos_tainty, panel_produccion,
   gestion_tainty y dashboard. Maneja la hoja "NUEVO PEDIDO" más
   PAGOS, INSUMOS_ESTADO, CLIENTES y PRODUCTOS.

   Este es el contenido completo a pegar en el editor de Apps Script
   y desplegar como nueva versión.

   Cambio vs versión anterior:
   - getPrecioSugerido ahora considera CANTIDAD además de vendedor.
     Cascada:
       1. mismo producto + vendedor + cantidad → fuente='vendedor_cantidad'
       2. mismo producto + vendedor             → fuente='vendedor'
       3. mismo producto (cualquier vendedor)   → fuente='generico'
     El cliente (formulario_tainty.html) ya envía &cantidad=N y
     muestra una etiqueta distinta según la fuente.
   ========================================================= */

const SHEET_ID = '1zIM2d9d-ILpChn4EVlLFkkZY_Rgod0y3i1a0dxWNatQ';
const SHEET_NAME = 'NUEVO PEDIDO';
const CLIENTES_SHEET = 'CLIENTES';
const PRODUCTOS_SHEET = 'PRODUCTOS';

function doPost(e) {
  try {
    if (e.parameter && e.parameter.action === 'setEstado') {
      const ok = setEstadoInsumo(e.parameter.pedido, e.parameter.insumo, e.parameter.estado);
      return buildResponse({ success: ok });
    }
    if (e.parameter && e.parameter.action === 'setEstadoTicket') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName(SHEET_NAME);
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === e.parameter.pedido) {
          sheet.getRange(i + 1, 14).setValue(e.parameter.estado);
          return buildResponse({ success: true });
        }
      }
      return buildResponse({ success: false, error: 'Pedido no encontrado' });
    }
    if (e.parameter && e.parameter.action === 'registrarAbono') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('PAGOS');
      const data = sheet.getDataRange().getValues();
      const num = parseInt(e.parameter.numAbono);
      const monto = parseFloat(e.parameter.monto);
      const fecha = e.parameter.fecha;
      const colAbono = num === 1 ? 7 : num === 2 ? 10 : num === 3 ? 13 : 16;
      const colFecha = colAbono + 1;
      const colRestante = colAbono + 2;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === e.parameter.pedido) {
          const total = parseFloat(data[i][5]) || 0;
          const prevAbonos = [data[i][6],data[i][9],data[i][12],data[i][15]].map(a=>parseFloat(a)||0).reduce((s,a)=>s+a,0);
          const nuevoRestante = Math.max(0, total - prevAbonos - monto);
          sheet.getRange(i+1, colAbono).setValue(monto);
          sheet.getRange(i+1, colFecha).setValue(fecha);
          sheet.getRange(i+1, colRestante).setValue(nuevoRestante);
          if (nuevoRestante <= 0) sheet.getRange(i+1, 19).setValue('PAGO REALIZADO');
          return buildResponse({ success: true });
        }
      }
      return buildResponse({ success: false, error: 'Pedido no encontrado' });
    }

    // Catálogos
    if (e.parameter && e.parameter.action === 'agregarCliente') {
      return buildResponse(agregarASimple(CLIENTES_SHEET, e.parameter.nombre));
    }
    if (e.parameter && e.parameter.action === 'agregarProducto') {
      return buildResponse(agregarASimple(PRODUCTOS_SHEET, e.parameter.nombre));
    }
    if (e.parameter && e.parameter.action === 'editarCliente') {
      return buildResponse(editarItemCatalogo(CLIENTES_SHEET, e.parameter.viejo, e.parameter.nuevo, 'cliente'));
    }
    if (e.parameter && e.parameter.action === 'editarProducto') {
      return buildResponse(editarItemCatalogo(PRODUCTOS_SHEET, e.parameter.viejo, e.parameter.nuevo, 'producto'));
    }
    if (e.parameter && e.parameter.action === 'eliminarCliente') {
      return buildResponse(eliminarItemCatalogo(CLIENTES_SHEET, e.parameter.nombre));
    }
    if (e.parameter && e.parameter.action === 'eliminarProducto') {
      return buildResponse(eliminarItemCatalogo(PRODUCTOS_SHEET, e.parameter.nombre));
    }

    let data;
    if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else {
      data = JSON.parse(e.postData.contents);
    }
    return buildResponse(guardarNuevoPedido(data));
  } catch(err) {
    return buildResponse({ success: false, error: err.message });
  }
}

// Lógica compartida: guarda un pedido (usado por doPost y doGet)
function guardarNuevoPedido(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const colA = sheet.getRange('A:A').getValues();
  let nuevaFila = 2;
  for (let i = 1; i < colA.length; i++) {
    if (colA[i][0] === '') { nuevaFila = i + 1; break; }
  }
  const filaAnterior = sheet.getRange(nuevaFila - 1, 1).getValue().toString();
  const ultimoNum = parseInt(filaAnterior.replace('P', '')) || 211;
  const nuevoPedido = 'P' + String(ultimoNum + 1).padStart(4, '0');
  const fechaBase = new Date(data.fecha + 'T12:00:00');
  const fechaFormateada = Utilities.formatDate(fechaBase, 'America/Lima', 'dd/MM/yyyy');
  const cantidad = parseInt(data.cantidad);
  const precio = parseFloat(data.precio);
  const abono = parseFloat(data.abono) || 0;
  const montoTotal = cantidad * precio;
  const pctAbonado = montoTotal > 0 ? Math.round((abono / montoTotal) * 100) + '%' : '0%';
  const meses = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];
  const mes = meses[fechaBase.getMonth()];
  const fechaTentativa = sumarDiasHabiles(fechaBase, 4);
  const fechaTentativaStr = Utilities.formatDate(fechaTentativa, 'America/Lima', 'dd/MM/yyyy');
  sheet.getRange(nuevaFila, 1).setValue(nuevoPedido);
  sheet.getRange(nuevaFila, 3).setValue(fechaFormateada);
  sheet.getRange(nuevaFila, 4).setValue(data.cliente);
  sheet.getRange(nuevaFila, 5).setValue(data.producto);
  sheet.getRange(nuevaFila, 6).setValue(cantidad);
  sheet.getRange(nuevaFila, 7).setValue(data.vendedor);
  sheet.getRange(nuevaFila, 8).setValue(precio);
  sheet.getRange(nuevaFila, 9).setValue(montoTotal);
  sheet.getRange(nuevaFila, 10).setValue(abono);
  sheet.getRange(nuevaFila, 11).setValue(pctAbonado);
  sheet.getRange(nuevaFila, 12).setValue(mes);
  sheet.getRange(nuevaFila, 13).setValue(fechaTentativaStr);
  sheet.getRange(nuevaFila, 15).setValue(data.observaciones || '');
  try { sheet.getRange(nuevaFila, 13).setNumberFormat('dd/MM/yyyy'); } catch(e) {}
  try { sheet.getRange(nuevaFila, 1, 1, 15).setHorizontalAlignment('center'); } catch(e) {}

  try { if (data.cliente) agregarASimple(CLIENTES_SHEET, data.cliente); } catch(e) {}
  try { if (data.producto) agregarASimple(PRODUCTOS_SHEET, data.producto); } catch(e) {}

  SpreadsheetApp.flush();
  return { success: true, pedido: nuevoPedido };
}

function sumarDiasHabiles(fecha, dias) {
  const result = new Date(fecha);
  let sumados = 0;
  while (sumados < dias) {
    result.setDate(result.getDate() + 1);
    if (result.getDay() !== 0) sumados++;
  }
  return result;
}

function doGet(e) {
  let result;
  if (e && e.parameter && e.parameter.action === 'precio') {
    // CAMBIO: ahora pasa también cantidad para que la sugerencia
    // priorice el último precio para esa misma escala/cantidad.
    result = getPrecioSugerido(e.parameter.producto, e.parameter.vendedor, e.parameter.cantidad);
  } else if (e && e.parameter && e.parameter.action === 'guardarPedido') {
    try {
      result = guardarNuevoPedido(e.parameter);
    } catch(err) {
      result = { success: false, error: err.message };
    }
  } else if (e && e.parameter && e.parameter.action === 'ultimo') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const ultimo = lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : 'P0211';
    result = { ultimo: ultimo };
  } else if (e && e.parameter && e.parameter.action === 'panel') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const pedidos = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const estado = (row[13] || '').toString().toLowerCase();
      if (estado.includes('entregado')) continue;
      pedidos.push({
        numPedido: row[0], fecha: row[2], cliente: row[3],
        producto: row[4], cantidad: row[5], vendedor: row[6],
        precio: row[7], montoTotal: row[8], abono: row[9],
        fechaTentativa: row[12], estado: row[13], observaciones: row[14]
      });
    }
    result = { success: true, pedidos: pedidos };
  } else if (e && e.parameter && e.parameter.action === 'estados') {
    result = { success: true, estados: getEstadosInsumos() };
  } else if (e && e.parameter && e.parameter.action === 'analytics') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      rows.push({
        numPedido: row[0],
        fecha: row[2],
        cliente: row[3],
        producto: row[4],
        cantidad: row[5],
        precio: row[7],
        montoTotal: row[8],
        mes: row[11],
        estado: row[13]
      });
    }
    result = { success: true, rows: rows };
  } else if (e && e.parameter && e.parameter.action === 'pagos') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('PAGOS');
    const data = sheet.getDataRange().getValues();
    const pagos = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0]) continue;
      pagos.push({
        numPedido:r[0], fecha:r[1], cliente:r[2],
        cantidad:r[3], producto:r[4], montoTotal:r[5],
        abono1:r[6], fecha1:r[7], restante1:r[8],
        abono2:r[9], fecha2:r[10], restante2:r[11],
        abono3:r[12], fecha3:r[13], restante3:r[14],
        abono4:r[15], fecha4:r[16], restante4:r[17],
        detalle:r[18], observaciones:r[19]
      });
    }
    result = { success: true, pagos: pagos };

  } else if (e && e.parameter && e.parameter.action === 'pedidos_activos') {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const pedidos = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const estado = (row[13] || '').toString().toLowerCase();
      if (estado.includes('entregado') || estado.includes('cancelado')) continue;
      pedidos.push({
        numPedido: row[0],
        fecha: row[2],
        cliente: row[3],
        producto: row[4],
        cantidad: row[5],
        estado: row[13]
      });
    }
    result = { success: true, pedidos: pedidos };

  } else if (e && e.parameter && e.parameter.action === 'clientes') {
    result = { success: true, data: leerListaSimple(CLIENTES_SHEET) };
  } else if (e && e.parameter && e.parameter.action === 'productos') {
    result = { success: true, data: leerListaSimple(PRODUCTOS_SHEET) };
  } else if (e && e.parameter && e.parameter.action === 'agregarCliente') {
    result = agregarASimple(CLIENTES_SHEET, e.parameter.nombre);
  } else if (e && e.parameter && e.parameter.action === 'agregarProducto') {
    result = agregarASimple(PRODUCTOS_SHEET, e.parameter.nombre);
  } else if (e && e.parameter && e.parameter.action === 'editarCliente') {
    result = editarItemCatalogo(CLIENTES_SHEET, e.parameter.viejo, e.parameter.nuevo, 'cliente');
  } else if (e && e.parameter && e.parameter.action === 'editarProducto') {
    result = editarItemCatalogo(PRODUCTOS_SHEET, e.parameter.viejo, e.parameter.nuevo, 'producto');
  } else if (e && e.parameter && e.parameter.action === 'eliminarCliente') {
    result = eliminarItemCatalogo(CLIENTES_SHEET, e.parameter.nombre);
  } else if (e && e.parameter && e.parameter.action === 'eliminarProducto') {
    result = eliminarItemCatalogo(PRODUCTOS_SHEET, e.parameter.nombre);
  } else if (e && e.parameter && e.parameter.action === 'inicializarCatalogos') {
    result = inicializarCatalogos();

  } else {
    result = { status: 'OK', message: 'Script activo' };
  }
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    const output = ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')');
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return output;
  }
  return buildResponse(result);
}

function buildResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Devuelve el último precio unitario sugerido para un producto, intentando
 * ser lo más específico posible.
 *
 * Estrategia (cascada de la más específica a la más general):
 *   1. Mismo producto + mismo vendedor + MISMA cantidad → 'vendedor_cantidad'
 *      (porque el precio por unidad cambia con la escala: 100 vs 500 vs 1000)
 *   2. Mismo producto + mismo vendedor (cualquier cantidad) → 'vendedor'
 *   3. Mismo producto (cualquier vendedor)                 → 'generico'
 *
 * Recorremos las filas de la última a la primera para que "último precio"
 * sea efectivamente el más reciente.
 */
/** Normaliza un string para comparar nombres tolerante a casing/acentos/espacios. */
function normNombre_(s) {
  if (s === null || s === undefined) return '';
  let v = s.toString().trim().toUpperCase();
  if (v.normalize) v = v.normalize('NFD').replace(/[̀-ͯ]/g, '');
  return v.replace(/\s+/g, ' ');
}

function getPrecioSugerido(nombreProducto, vendedor, cantidad) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  const prodNorm = normNombre_(nombreProducto);
  const vendNorm = vendedor ? normNombre_(vendedor) : null;
  const cantNum = parseInt(cantidad) || 0;

  let precioVendedorCantidad = null;
  let precioVendedor = null;
  let precioGenerico = null;

  for (let i = data.length - 1; i >= 1; i--) {
    // Comparación case-insensitive + sin acentos para que matchee aunque
    // el producto haya quedado guardado con distinto casing en pedidos viejos.
    if (normNombre_(data[i][4]) !== prodNorm) continue;
    const precio = parseFloat(data[i][7]) || 0;
    if (precio <= 0) continue;

    // Nivel 3 (cualquier vendedor) — siempre se toma el primero encontrado
    if (precioGenerico === null) precioGenerico = precio;

    if (vendNorm && data[i][6]) {
      const v = normNombre_(data[i][6]);
      if (v === vendNorm) {
        // Nivel 2 (vendedor, cualquier cantidad)
        if (precioVendedor === null) precioVendedor = precio;
        // Nivel 1 (vendedor + cantidad exacta)
        if (cantNum > 0 && precioVendedorCantidad === null) {
          const filaCant = parseInt(data[i][5]) || 0;
          if (filaCant === cantNum) {
            precioVendedorCantidad = precio;
            break; // ya tenemos el más específico
          }
        }
      }
    }
  }

  if (precioVendedorCantidad !== null) {
    return { precio: precioVendedorCantidad, fuente: 'vendedor_cantidad', success: true };
  }
  if (precioVendedor !== null) {
    return { precio: precioVendedor, fuente: 'vendedor', success: true };
  }
  if (precioGenerico !== null) {
    return { precio: precioGenerico, fuente: 'generico', success: true };
  }
  return { precio: null, fuente: 'ninguno', success: true };
}

function getEstadosInsumos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('INSUMOS_ESTADO');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const estados = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0] + '|' + data[i][1];
    estados[key] = data[i][2];
  }
  return estados;
}

function setEstadoInsumo(pedidoId, insumo, estado) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('INSUMOS_ESTADO');
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pedidoId && data[i][1] === insumo) {
      sheet.getRange(i + 1, 3).setValue(estado);
      sheet.getRange(i + 1, 4).setValue(new Date());
      return true;
    }
  }
  sheet.appendRow([pedidoId, insumo, estado, new Date()]);
  return true;
}

// =================== CATÁLOGOS ===================

function leerListaSimple(nombreHoja) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) sheet = crearHojaCatalogo(ss, nombreHoja);
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const v = (data[i][0] || '').toString().trim();
    if (v) out.push(v);
  }
  return Array.from(new Set(out)).sort((a, b) => a.localeCompare(b, 'es'));
}

function agregarASimple(nombreHoja, valor) {
  try {
    if (!valor) return { success: false, error: 'Valor vacío' };
    const v = valor.toString().trim();
    if (!v) return { success: false, error: 'Valor vacío' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) sheet = crearHojaCatalogo(ss, nombreHoja);

    const data = sheet.getDataRange().getValues();
    const vNorm = v.toLowerCase();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim().toLowerCase() === vNorm) {
        return { success: true, agregado: false, motivo: 'Ya existía' };
      }
    }
    const fechaStr = Utilities.formatDate(new Date(), 'America/Lima', 'dd/MM/yyyy HH:mm');
    sheet.appendRow([v, fechaStr]);
    return { success: true, agregado: true, valor: v };
  } catch (err) {
    return { success: false, error: 'agregarASimple: ' + err.message };
  }
}

function editarItemCatalogo(nombreHoja, viejo, nuevo, tipo) {
  try {
    if (!viejo || !nuevo) return { success: false, error: 'Faltan parámetros' };
    const viejoT = viejo.toString().trim();
    const nuevoT = nuevo.toString().trim();
    if (!viejoT || !nuevoT) return { success: false, error: 'Valor vacío' };
    if (viejoT.toLowerCase() === nuevoT.toLowerCase()) {
      return { success: true, motivo: 'Sin cambios' };
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return { success: false, error: 'Hoja ' + nombreHoja + ' no existe' };

    const data = sheet.getDataRange().getValues();
    const nuevoNorm = nuevoT.toLowerCase();
    const viejoNorm = viejoT.toLowerCase();
    let filaAEditar = -1;
    for (let i = 1; i < data.length; i++) {
      const v = (data[i][0] || '').toString().trim().toLowerCase();
      if (v === nuevoNorm && v !== viejoNorm) {
        return { success: false, error: 'Ya existe un item con ese nombre nuevo' };
      }
      if (v === viejoNorm) filaAEditar = i + 1;
    }
    if (filaAEditar < 0) return { success: false, error: 'No se encontró el item original' };

    sheet.getRange(filaAEditar, 1).setValue(nuevoT);

    const sheetPed = ss.getSheetByName(SHEET_NAME);
    const colPed = tipo === 'cliente' ? 4 : 5;
    let actualizados = 0;
    if (sheetPed) {
      const dataPed = sheetPed.getDataRange().getValues();
      for (let i = 1; i < dataPed.length; i++) {
        const valActual = (dataPed[i][colPed - 1] || '').toString().trim();
        if (valActual.toLowerCase() === viejoNorm) {
          sheetPed.getRange(i + 1, colPed).setValue(nuevoT);
          actualizados++;
        }
      }
    }

    SpreadsheetApp.flush();
    return { success: true, viejo: viejoT, nuevo: nuevoT, pedidosActualizados: actualizados };
  } catch (err) {
    return { success: false, error: 'editarItemCatalogo: ' + err.message };
  }
}

function eliminarItemCatalogo(nombreHoja, valor) {
  try {
    if (!valor) return { success: false, error: 'Valor vacío' };
    const v = valor.toString().trim();
    if (!v) return { success: false, error: 'Valor vacío' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) return { success: false, error: 'Hoja ' + nombreHoja + ' no existe' };

    const data = sheet.getDataRange().getValues();
    const vNorm = v.toLowerCase();
    let borradas = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      const cell = (data[i][0] || '').toString().trim().toLowerCase();
      if (cell === vNorm) {
        sheet.deleteRow(i + 1);
        borradas++;
      }
    }
    if (borradas === 0) return { success: false, error: 'No se encontró el item' };

    SpreadsheetApp.flush();
    return { success: true, eliminado: v, filas: borradas };
  } catch (err) {
    return { success: false, error: 'eliminarItemCatalogo: ' + err.message };
  }
}

function crearHojaCatalogo(ss, nombre) {
  const sheet = ss.insertSheet(nombre);
  sheet.getRange(1, 1, 1, 2).setValues([['Nombre', 'Fecha registro']]);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  return sheet;
}

function inicializarCatalogos() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const creadas = [];
  if (!ss.getSheetByName(CLIENTES_SHEET)) { crearHojaCatalogo(ss, CLIENTES_SHEET); creadas.push(CLIENTES_SHEET); }
  if (!ss.getSheetByName(PRODUCTOS_SHEET)) { crearHojaCatalogo(ss, PRODUCTOS_SHEET); creadas.push(PRODUCTOS_SHEET); }

  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const clientesUnicos = new Set();
  const productosUnicos = new Set();
  for (let i = 1; i < data.length; i++) {
    const c = (data[i][3] || '').toString().trim();
    const p = (data[i][4] || '').toString().trim();
    if (c) clientesUnicos.add(c);
    if (p) productosUnicos.add(p);
  }
  let cAgregados = 0, pAgregados = 0;
  Array.from(clientesUnicos).forEach(c => { if (agregarASimple(CLIENTES_SHEET, c).agregado) cAgregados++; });
  Array.from(productosUnicos).forEach(p => { if (agregarASimple(PRODUCTOS_SHEET, p).agregado) pAgregados++; });

  return {
    success: true,
    hojasCreadas: creadas,
    clientesAgregados: cAgregados,
    productosAgregados: pAgregados
  };
}
