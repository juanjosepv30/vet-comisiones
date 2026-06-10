function _rows(sheet) {
  const last = sheet.getLastRow();
  if (last <= 1) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues()
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = r[i] instanceof Date
          ? Utilities.formatDate(r[i], 'America/Bogota', 'yyyy-MM-dd')
          : r[i];
      });
      return obj;
    });
}

function _nextId(sheet) {
  const last = sheet.getLastRow();
  if (last <= 1) return 1;
  return Number(sheet.getRange(last, 1).getValue() || 0) + 1;
}

function _rowFor(sheet, id) {
  const last = sheet.getLastRow();
  if (last <= 1) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const i = ids.indexOf(Number(id));
  return i === -1 ? -1 : i + 2;
}

// ── Empleados ─────────────────────────────────────────────────────────────────

function getEmpleados() {
  return _rows(getSheet('Empleados'));
}

function saveEmpleado(data) {
  const s = getSheet('Empleados');
  const row = [data.id ? Number(data.id) : _nextId(s), data.nombre, data.cargo, Number(data.comision)];
  if (data.id) {
    const r = _rowFor(s, data.id);
    if (r < 0) throw new Error('Empleado no encontrado');
    s.getRange(r, 1, 1, 4).setValues([row]);
  } else {
    s.appendRow(row);
  }
  return { ok: true };
}

function deleteEmpleado(id) {
  const s = getSheet('Empleados');
  const r = _rowFor(s, id);
  if (r < 0) throw new Error('Empleado no encontrado');
  s.deleteRow(r);
  return { ok: true };
}

// ── Servicios ─────────────────────────────────────────────────────────────────

function getServicios() {
  return _rows(getSheet('Servicios'));
}

function saveServicio(data) {
  const s = getSheet('Servicios');
  const row = [data.id ? Number(data.id) : _nextId(s), data.nombre, Number(data.precio), Number(data.comision)];
  if (data.id) {
    const r = _rowFor(s, data.id);
    if (r < 0) throw new Error('Servicio no encontrado');
    s.getRange(r, 1, 1, 4).setValues([row]);
  } else {
    s.appendRow(row);
  }
  return { ok: true };
}

function deleteServicio(id) {
  const s = getSheet('Servicios');
  const r = _rowFor(s, id);
  if (r < 0) throw new Error('Servicio no encontrado');
  s.deleteRow(r);
  return { ok: true };
}

// ── Registros ─────────────────────────────────────────────────────────────────

function addRegistro(data) {
  const s = getSheet('Registros');
  const id = _nextId(s);
  const monto = Math.round((Number(data.precio) * Number(data.comision)) / 100);
  s.appendRow([id, data.fecha, data.empleado, data.servicio, Number(data.precio), Number(data.comision), monto, 'Pendiente']);
  return { ok: true, id };
}

function getRegistros(filtros) {
  filtros = filtros || {};
  return _rows(getSheet('Registros')).filter(r => {
    if (filtros.empleado && r.Empleado !== filtros.empleado) return false;
    if (filtros.estado  && r.Estado   !== filtros.estado)   return false;
    if (filtros.desde   && r.Fecha    <  filtros.desde)     return false;
    if (filtros.hasta   && r.Fecha    >  filtros.hasta)     return false;
    return true;
  });
}

function getResumen(desde, hasta) {
  const regs = getRegistros({ desde, hasta });
  const map = {};
  regs.forEach(r => {
    if (!map[r.Empleado]) map[r.Empleado] = { empleado: r.Empleado, ventas: 0, comision: 0, registros: 0 };
    map[r.Empleado].ventas    += Number(r.Precio) || 0;
    map[r.Empleado].comision  += Number(r.Monto)  || 0;
    map[r.Empleado].registros += 1;
  });
  return Object.values(map).sort((a, b) => b.comision - a.comision);
}

function marcarPagada(id) {
  const s = getSheet('Registros');
  const r = _rowFor(s, id);
  if (r < 0) throw new Error('Registro no encontrado');
  s.getRange(r, 8).setValue('Pagada');
  return { ok: true };
}
