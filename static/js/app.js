// ═══════════════════════════════════════════════
//  ESTADO GLOBAL
// ═══════════════════════════════════════════════
let estado = {
  datos: null,        // resultado del servidor
  modalCtx: null,     // {empIdx, diaIdx}
};

// ═══════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  // Poblar selector de años
  const sel = document.getElementById('anio');
  const anioActual = new Date().getFullYear();
  for (let a = anioActual; a >= 2016; a--) {
    const opt = document.createElement('option');
    opt.value = a; opt.textContent = a;
    sel.appendChild(opt);
  }
  // Mes actual por defecto
  document.getElementById('mes').value = new Date().getMonth() + 1;
});

// ═══════════════════════════════════════════════
//  ARCHIVOS
// ═══════════════════════════════════════════════
function onFileSelect(input, tipo) {
  const file = input.files[0];
  if (!file) return;
  const drop = document.getElementById(`drop-${tipo}`);
  const name = document.getElementById(`name-${tipo}`);
  drop.classList.add('has-file');
  name.textContent = '✓ ' + file.name;
}

// ═══════════════════════════════════════════════
//  PROCESAR
// ═══════════════════════════════════════════════
async function procesarQuincena() {
  const reloj    = document.getElementById('file-reloj').files[0];
  const personal = document.getElementById('file-personal').files[0];
  if (!reloj || !personal) {
    alert('Por favor cargá ambos archivos antes de continuar.');
    return;
  }

  const fd = new FormData();
  fd.append('reloj',    reloj);
  fd.append('personal', personal);
  fd.append('quincena', document.getElementById('quincena').value);
  fd.append('anio',     document.getElementById('anio').value);
  fd.append('mes',      document.getElementById('mes').value);
  fd.append('sector',   document.getElementById('sector').value);
  fd.append('lista_negro', document.getElementById('lista-negro').value);

  mostrarLoading(true);
  try {
    const res  = await fetch('/api/procesar', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || 'Error del servidor');

    estado.datos = data;
    renderTabla(data);
    document.getElementById('paso-config').classList.add('hidden');
    document.getElementById('paso-tabla').classList.remove('hidden');
  } catch (err) {
    alert('Error al procesar: ' + err.message);
  } finally {
    mostrarLoading(false);
  }
}

function volverConfig() {
  document.getElementById('paso-tabla').classList.add('hidden');
  document.getElementById('paso-config').classList.remove('hidden');
}

// ═══════════════════════════════════════════════
//  RENDER TABLA
// ═══════════════════════════════════════════════
function renderTabla(data) {
  // Título
  const q = data.quincena === 1 ? '1ª Quincena (1–15)' : `2ª Quincena (16–${data.fecha_fin.slice(8,10)})`;
  document.getElementById('tabla-titulo').textContent =
    `${data.sector} — ${data.mes_nombre} ${data.anio}`;
  document.getElementById('tabla-subtitulo').textContent = q;

  // Alertas mapeo
  const alertDiv = document.getElementById('alertas-mapeo');
  if (data.sin_mapeo && data.sin_mapeo.length > 0) {
    alertDiv.classList.remove('hidden');
    alertDiv.innerHTML = `<strong>⚠ ${data.sin_mapeo.length} empleado(s) sin registros en el reloj:</strong>
      Sus horas figuran en cero. Podés editarlas manualmente haciendo clic en cada celda.`;
  } else {
    alertDiv.classList.add('hidden');
  }

  // THEAD
  const thead = document.getElementById('tabla-thead');
  thead.innerHTML = '';
  const trH = document.createElement('tr');

  const colsFijas = ['#', 'Legajo', 'Nombre', '12hs'];
  colsFijas.forEach((t, i) => {
    const th = document.createElement('th');
    th.textContent = t;
    if (t === 'Nombre') th.classList.add('col-nombre');
    trH.appendChild(th);
  });

  data.dias.forEach(dia => {
    const th = document.createElement('th');
    th.classList.add('col-dia');
    const tipo = dia.tipo;
    if (tipo === 'feriado') th.classList.add('col-feriado');
    if (tipo === 'sabado')  th.classList.add('col-sabado');
    const dow = ['L','M','X','J','V','S','D'][dia.dow];
    th.innerHTML = `<div>${dow}</div><div>${dia.label}</div>`;
    th.title = `${dia.label} — ${tipo}`;
    trH.appendChild(th);
  });

  // Columnas totales
  ['N', 'E50', 'E100', 'Enf', 'Acc', 'Total'].forEach(t => {
    const th = document.createElement('th');
    th.textContent = t;
    th.title = {N:'Hs Normales', E50:'Extras 50%', E100:'Extras 100%',
                Enf:'Enfermedad', Acc:'Accidente', Total:'Total Hs'}[t];
    trH.appendChild(th);
  });

  thead.appendChild(trH);

  // TBODY
  const tbody = document.getElementById('tabla-tbody');
  tbody.innerHTML = '';

  data.empleados.forEach((emp, empIdx) => {
    const tr = document.createElement('tr');
    if (emp.sin_datos)  tr.classList.add('fila-sin-datos');
    if (emp.dias.some(d => d.estado === 'incompleto')) tr.classList.add('fila-incompleto');

    // N°
    tdFijo(tr, empIdx + 1, 'col-legajo');
    // Legajo
    tdFijo(tr, emp.legajo, 'col-legajo');
    // Nombre
    tdFijo(tr, emp.nombre, 'col-nombre');
    // Toggle 12hs
    const tdJ = document.createElement('td');
    tdJ.classList.add('col-jornada12');
    const chk = document.createElement('input');
    chk.type = 'checkbox';
    chk.checked = emp.jornada_12;
    chk.addEventListener('change', () => {
      emp.jornada_12 = chk.checked;
      recalcularEmpleado(emp);
      actualizarFilaTotales(tr, emp, empIdx);
    });
    tdJ.appendChild(chk);
    tr.appendChild(tdJ);

    // Celdas de días
    emp.dias.forEach((dia, diaIdx) => {
      const td = document.createElement('td');
      const div = document.createElement('div');
      div.classList.add('celda-dia', dia.editado ? 'editado' : dia.estado);
      div.title = `${dia.fecha} — Click para editar`;
      div.textContent = resumenCelda(dia);
      div.addEventListener('click', () => abrirModal(empIdx, diaIdx));
      td.appendChild(div);
      tr.appendChild(td);
    });

    // Totales
    const t = emp.totales;
    const total = round2(t.hs_normales + t.hs_ext50 + t.hs_ext100 + t.hs_enfermedad + t.hs_accidente);
    [t.hs_normales, t.hs_ext50, t.hs_ext100, t.hs_enfermedad, t.hs_accidente, total].forEach((v, i) => {
      const td = document.createElement('td');
      td.classList.add('col-totales');
      if (emp.en_negro && i === 1) td.classList.add('col-negro');
      if (emp.en_negro && i === 2) td.classList.add('col-negro');
      td.textContent = v > 0 ? v : '—';
      td.id = `tot-${empIdx}-${i}`;
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

function tdFijo(tr, val, cls) {
  const td = document.createElement('td');
  if (cls) td.classList.add(cls);
  td.textContent = val;
  tr.appendChild(td);
}

function resumenCelda(dia) {
  if (dia.estado === 'descanso') return '—';
  if (dia.estado === 'ausente')  return 'AUS';
  if (dia.estado === 'incompleto') return '⚠';
  const norm = dia.hs_normales;
  const ext  = dia.hs_ext50 + dia.hs_ext100;
  if (norm === 0 && ext === 0) return '—';
  let txt = norm > 0 ? String(norm) : '';
  if (ext > 0) txt += (txt ? '+' : '') + ext + 'e';
  return txt;
}

// ═══════════════════════════════════════════════
//  MODAL EDICIÓN
// ═══════════════════════════════════════════════
function abrirModal(empIdx, diaIdx) {
  const emp = estado.datos.empleados[empIdx];
  const dia = emp.dias[diaIdx];
  estado.modalCtx = { empIdx, diaIdx };

  const TIPOS = { normal:'Día laboral', viernes:'Viernes', sabado:'Sábado',
                   feriado:'Feriado', domingo:'Domingo' };
  document.getElementById('modal-titulo').textContent =
    `${emp.nombre} — ${dia.fecha}`;
  document.getElementById('modal-info').textContent =
    `${TIPOS[dia.tipo_dia] || dia.tipo_dia}  |  Legajo ${emp.legajo}`;

  document.getElementById('modal-entrada').value    = dia.entrada || '';
  document.getElementById('modal-salida').value     = dia.salida  || '';
  document.getElementById('modal-normales').value   = dia.hs_normales   || 0;
  document.getElementById('modal-ext50').value      = dia.hs_ext50      || 0;
  document.getElementById('modal-ext100').value     = dia.hs_ext100     || 0;
  document.getElementById('modal-enfermedad').value = dia.hs_enfermedad || 0;
  document.getElementById('modal-accidente').value  = dia.hs_accidente  || 0;
  document.getElementById('modal-obs').value        = dia.obs || '';
  document.getElementById('modal-jornada12').checked = emp.jornada_12 || false;

  document.getElementById('modal-overlay').classList.remove('hidden');
  document.getElementById('modal-normales').focus();
}

function cerrarModal() {
  document.getElementById('modal-overlay').classList.add('hidden');
  estado.modalCtx = null;
}

function guardarModal() {
  const { empIdx, diaIdx } = estado.modalCtx;
  const emp = estado.datos.empleados[empIdx];
  const dia = emp.dias[diaIdx];

  dia.entrada       = document.getElementById('modal-entrada').value;
  dia.salida        = document.getElementById('modal-salida').value;
  dia.hs_normales   = parseFloat(document.getElementById('modal-normales').value)   || 0;
  dia.hs_ext50      = parseFloat(document.getElementById('modal-ext50').value)      || 0;
  dia.hs_ext100     = parseFloat(document.getElementById('modal-ext100').value)     || 0;
  dia.hs_enfermedad = parseFloat(document.getElementById('modal-enfermedad').value) || 0;
  dia.hs_accidente  = parseFloat(document.getElementById('modal-accidente').value)  || 0;
  dia.obs           = document.getElementById('modal-obs').value;
  emp.jornada_12    = document.getElementById('modal-jornada12').checked;

  dia.estado  = 'ok';
  dia.editado = true;
  if (dia.hs_normales === 0 && dia.hs_ext50 === 0 && dia.hs_ext100 === 0 &&
      dia.hs_enfermedad === 0 && dia.hs_accidente === 0) {
    dia.estado = 'ausente';
  }

  recalcularTotalesEmpleado(emp);
  renderTabla(estado.datos);
  cerrarModal();
}

function marcarAusente() {
  const { empIdx, diaIdx } = estado.modalCtx;
  const dia = estado.datos.empleados[empIdx].dias[diaIdx];
  dia.hs_normales = dia.hs_ext50 = dia.hs_ext100 =
  dia.hs_enfermedad = dia.hs_accidente = 0;
  dia.entrada = dia.salida = '';
  dia.estado = 'ausente';
  dia.editado = true;
  dia.obs = document.getElementById('modal-obs').value;
  recalcularTotalesEmpleado(estado.datos.empleados[empIdx]);
  renderTabla(estado.datos);
  cerrarModal();
}

// ═══════════════════════════════════════════════
//  RECALCULAR
// ═══════════════════════════════════════════════
function recalcularEmpleado(emp) {
  // Re-procesa el empleado con la nueva jornada (solo si no fue editado manualmente)
  emp.dias.forEach(dia => {
    if (dia.editado) return;
    // La lógica de recalculo con jornada 12 se aplica del lado servidor
    // Aquí simplemente marcamos que necesita refresh si se usa jornada 12
  });
  recalcularTotalesEmpleado(emp);
}

function recalcularTotalesEmpleado(emp) {
  let norm = 0, e50 = 0, e100 = 0, enf = 0, acc = 0;
  emp.dias.forEach(d => {
    norm += d.hs_normales   || 0;
    e50  += d.hs_ext50      || 0;
    e100 += d.hs_ext100     || 0;
    enf  += d.hs_enfermedad || 0;
    acc  += d.hs_accidente  || 0;
  });
  emp.totales = {
    hs_normales:   round2(norm),
    hs_ext50:      round2(e50),
    hs_ext100:     round2(e100),
    hs_enfermedad: round2(enf),
    hs_accidente:  round2(acc),
  };
}

function actualizarFilaTotales(tr, emp, empIdx) {
  const t = emp.totales;
  const total = round2(t.hs_normales+t.hs_ext50+t.hs_ext100+t.hs_enfermedad+t.hs_accidente);
  [t.hs_normales, t.hs_ext50, t.hs_ext100, t.hs_enfermedad, t.hs_accidente, total].forEach((v, i) => {
    const td = document.getElementById(`tot-${empIdx}-${i}`);
    if (td) td.textContent = v > 0 ? v : '—';
  });
}

// ═══════════════════════════════════════════════
//  DESCARGAR
// ═══════════════════════════════════════════════
async function descargarExcel() {
  if (!estado.datos) return;
  mostrarLoading(true);
  try {
    const payload = {
      empleados: estado.datos.empleados,
      meta: {
        quincena:   estado.datos.quincena,
        anio:       estado.datos.anio,
        mes:        estado.datos.mes,
        sector:     estado.datos.sector,
        fecha_ini:  estado.datos.fecha_ini,
        fecha_fin:  estado.datos.fecha_fin,
      }
    };
    const res = await fetch('/api/descargar', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error('Error al generar Excel');
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url;
    a.download = res.headers.get('Content-Disposition')?.split('filename=')[1] ||
                 `Quincena_Q${estado.datos.quincena}_${estado.datos.mes_nombre}_${estado.datos.anio}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch (err) {
    alert('Error al descargar: ' + err.message);
  } finally {
    mostrarLoading(false);
  }
}

// ═══════════════════════════════════════════════
//  UTILS
// ═══════════════════════════════════════════════
function round2(n) { return Math.round(n * 100) / 100; }

function mostrarLoading(show) {
  document.getElementById('loading').classList.toggle('hidden', !show);
}
