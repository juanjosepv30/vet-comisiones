function ensureSetup() {
  const ss = getSpreadsheet();
  _initSheet(ss, 'Empleados',  ['ID', 'Nombre', 'Cargo', 'Comision']);
  _initSheet(ss, 'Servicios',  ['ID', 'Nombre', 'Precio', 'Comision']);
  _initSheet(ss, 'Registros',  ['ID', 'Fecha', 'Empleado', 'Servicio', 'Precio', 'Comision', 'Monto', 'Estado']);
  return { ok: true };
}

function _initSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const range = sheet.getRange(1, 1, 1, headers.length);
    range.setValues([headers]);
    range.setBackground('#2c7a4b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}
