/* =========================================================
   SNIPPET — action=precio en el SCRIPT_URL (el de pedidos/pagos)
   ---------------------------------------------------------
   Este archivo NO se despliega solo. Es el handler que el cliente
   (formulario_tainty.html) llama vía:
     SCRIPT_URL?action=precio&producto=X&vendedor=Y&cantidad=Z
   con el cambio nuevo en sugerirPrecio().
   ---------------------------------------------------------
   Si tu script SCRIPT_URL ya tiene un caso `action === 'precio'`
   en doGet, REEMPLAZÁ ese caso por la función handlerPrecio_ de
   abajo. Si nunca existió, agregalo al switch/if de doGet.

   Estrategia de búsqueda (cascada, devuelve el más específico):
     1. Mismo producto + vendedor + cantidad  → fuente='vendedor_cantidad'
     2. Mismo producto + vendedor             → fuente='vendedor'
     3. Mismo producto (cualquier vendedor)   → fuente='otro_vendedor'

   Importante: ajustá las constantes de columnas (COL_*) según las
   posiciones reales en tu hoja de pedidos. Mirá el orden de las
   columnas del sheet y poné el índice 0-based de cada una.
   ========================================================= */

// === CONFIG: ajustar según tu hoja "PEDIDOS" ===
const PEDIDOS_SHEET_ID = 'TU_SPREADSHEET_ID';   // ← reemplazar
const PEDIDOS_SHEET_NAME = 'PEDIDOS';           // ← nombre de la pestaña
const COL_PRODUCTO  = 3;   // ← columna del producto (0-based)
const COL_VENDEDOR  = 4;   // ← columna del vendedor
const COL_CANTIDAD  = 5;   // ← columna de la cantidad
const COL_PRECIO    = 6;   // ← columna del precio unitario
const COL_FECHA     = 1;   // ← columna de la fecha (opcional, para "último")

/**
 * Handler completo de action=precio.
 * Llamalo desde doGet cuando e.parameter.action === 'precio'.
 */
function handlerPrecio_(e) {
  const callback = e && e.parameter && e.parameter.callback;
  try {
    const producto = (e.parameter.producto || '').toString().trim();
    const vendedor = (e.parameter.vendedor || '').toString().trim();
    const cantidad = parseInt(e.parameter.cantidad) || 0;
    if (!producto || !vendedor) {
      return respuestaJsonp_(callback, { success: false, error: 'Falta producto o vendedor' });
    }

    const sheet = SpreadsheetApp.openById(PEDIDOS_SHEET_ID).getSheetByName(PEDIDOS_SHEET_NAME);
    if (!sheet) {
      return respuestaJsonp_(callback, { success: false, error: 'Hoja no encontrada: ' + PEDIDOS_SHEET_NAME });
    }
    const data = sheet.getDataRange().getValues();

    const normN = s => (s || '').toString().trim().toUpperCase();
    const prodN = normN(producto);
    const vendN = normN(vendedor);

    // Recorrer de la última fila hacia arriba para encontrar el "último" rápido
    let porVendedorYCantidad = null;
    let porVendedor = null;
    let porProductoSolo = null;

    for (let i = data.length - 1; i >= 1; i--) {  // i=1 → salta cabecera
      const r = data[i];
      const pName = normN(r[COL_PRODUCTO]);
      if (pName !== prodN) continue;

      const precio = parseFloat(r[COL_PRECIO]) || 0;
      if (precio <= 0) continue;

      const vName = normN(r[COL_VENDEDOR]);
      const cant = parseInt(r[COL_CANTIDAD]) || 0;

      // Nivel 1: exact match vendedor + cantidad
      if (porVendedorYCantidad === null && vName === vendN && cant === cantidad && cantidad > 0) {
        porVendedorYCantidad = precio;
      }
      // Nivel 2: vendedor cualquier cantidad
      if (porVendedor === null && vName === vendN) {
        porVendedor = precio;
      }
      // Nivel 3: cualquier vendedor
      if (porProductoSolo === null) {
        porProductoSolo = precio;
      }
      // Si ya tenemos el más específico, podemos cortar
      if (porVendedorYCantidad !== null) break;
    }

    if (porVendedorYCantidad !== null) {
      return respuestaJsonp_(callback, { success: true, precio: porVendedorYCantidad, fuente: 'vendedor_cantidad' });
    }
    if (porVendedor !== null) {
      return respuestaJsonp_(callback, { success: true, precio: porVendedor, fuente: 'vendedor' });
    }
    if (porProductoSolo !== null) {
      return respuestaJsonp_(callback, { success: true, precio: porProductoSolo, fuente: 'otro_vendedor' });
    }
    return respuestaJsonp_(callback, { success: true, precio: 0, fuente: null });
  } catch (err) {
    console.error('handlerPrecio_ error:', err.message, err.stack);
    return respuestaJsonp_(callback, { success: false, error: err.message, stack: err.stack });
  }
}

/** Helper local — copia/pega o reusá el tuyo. */
function respuestaJsonp_(callback, obj) {
  const result = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + result + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(result)
    .setMimeType(ContentService.MimeType.JSON);
}
