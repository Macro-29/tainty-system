/* =========================================================
   Tainty — config.js
   Centralized Google Apps Script endpoints.
   Update here when redeploying the Apps Script.
   ========================================================= */

// Main Apps Script: pedidos, clientes, productos, pagos, panel, estados, analytics
const SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx1Nh7OzKW8p5D2Xp4iKrVnicIKxZaSUDkaFHxrGAUHZKJ1kRlWmOgOUNeNTzqaSTrWdg/exec';

// Costeo Apps Script: catalogo de costos y precios
const COSTEO_URL = 'https://script.google.com/macros/s/AKfycbyZYFJbmxl82uq-SR66LCmxuXhh_PkGpNAaspdjzYBTrnO0TckU90ZDyzGzXjtRtdwQ/exec';

// Pedidos activos (mismo endpoint que SCRIPT_URL, alias para claridad)
const PEDIDOS_URL = SCRIPT_URL;
