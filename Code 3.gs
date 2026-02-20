// =====================================================
// Trayectorias Secundaria - Backend (Supabase + Apps Script)
// - Mantiene el MISMO contrato de API que tu frontend ya usa.
// - Ya NO usa Google Sheets (no existe sheet_()).
// - Datos en Supabase (Postgres) vía REST.
// =====================================================

// Tablas en Supabase (nombres exactos)
const SHEETS = {
  ESTUDIANTES: 'estudiantes',
  CATALOGO: 'materias_catalogo',
  ESTADO: 'estado_por_ciclo',
  AUDITORIA: 'auditoria',
  EGRESADOS: 'egresados',
  APROBADAS_LIMPIEZA: 'materias_aprobadas_limpieza'
};

// Script Properties (Apps Script → Project Settings → Script properties)
const PROP_API_KEY = 'TRAYECTORIAS_API_KEY';
const PROP_SUPABASE_URL = 'SUPABASE_URL';
const PROP_SUPABASE_SERVICE_KEY = 'SUPABASE_SERVICE_KEY';

// Column order fijo (emula headers/rows como Sheets)
const TABLE_COLS = {
  estudiantes: ['id_estudiante','dni','apellido','nombre','anio_actual','division','turno','activo','observaciones','orientacion','egresado','anio_egreso','ciclo_egreso','fecha_pase_egresados'],
  materias_catalogo: ['id_materia','nombre','anio','es_troncal','orientacion','egresado','anio_egreso'],
  estado_por_ciclo: ['ciclo_lectivo','id_estudiante','id_materia','condicion_academica','nunca_cursada','situacion_actual','motivo_no_cursa','fecha_actualizacion','usuario','resultado_cierre','ciclo_cerrado'],
  auditoria: ['timestamp','ciclo_lectivo','id_estudiante','id_materia','campo','antes','despues','usuario'],
  egresados: ['id_estudiante','apellido','nombre','division','turno','ciclo_egreso','fecha_pase_egresados','observaciones'],
  materias_aprobadas_limpieza: ['ciclo_lectivo','id_estudiante','id_materia','condicion_academica','nunca_cursada','situacion_actual','motivo_no_cursa','fecha_actualizacion','usuario','resultado_cierre','ciclo_cerrado']
};

// ======== Menú (opcional, si está vinculado a una Sheet) ========
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('📘 Trayectorias (Supabase)')
      .addItem('🔑 Generar/Mostrar API Key', 'setupApiKey')
      .addItem('🧪 Probar API (ping)', 'testPing')
      .addToUi();
  } catch (err) {}
}

function setupApiKey() {
  const props = PropertiesService.getScriptProperties();
  let key = props.getProperty(PROP_API_KEY);
  if (!key) {
    key = Utilities.getUuid();
    props.setProperty(PROP_API_KEY, key);
  }
  try {
    SpreadsheetApp.getUi().alert('API Key (guardala):\n\n' + key);
  } catch (err) {
    Logger.log('API Key: ' + key);
  }
  return key;
}

function testPing() {
  const key = PropertiesService.getScriptProperties().getProperty(PROP_API_KEY);
  const res = handleRequest_({ apiKey: key, action: 'ping', payload: {} });
  try { SpreadsheetApp.getUi().alert(JSON.stringify(res, null, 2)); } catch(e) { Logger.log(JSON.stringify(res)); }
}

// ======== Web App entrypoint ========
function doPost(e) {
  try {
    const body = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    const req = body ? JSON.parse(body) : {};
    const result = handleRequest_(req);
    return jsonOut_(result, 200);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: (err && err.stack) ? String(err.stack) : null }, 500);
  }
}

// GET para test sin consola:
// /exec?action=getCycles&apiKey=XXX
// /exec?action=getStudentList&apiKey=XXX&ciclo_lectivo=2026
function doGet(e) {
  const p = (e && e.parameter) ? e.parameter : {};
  const action = String(p.action || '').trim();

  if (!action) {
    return jsonOut_({
      ok: true,
      service: 'Trayectorias Backend (Supabase)',
      endpoints: ['POST {apiKey, action, payload}', 'GET /exec?action=...&apiKey=...'],
      actions: ['ping','diag','getCycles','getCatalog','getStudentList','getStudentStatus','saveStudentStatus','syncCatalogRows','rolloverCycle','getDivisionRiskSummary','closeCycle','updateStudentOrientation'],
      examples: [
        '?action=getCycles&apiKey=TU_KEY',
        '?action=getStudentList&apiKey=TU_KEY&ciclo_lectivo=2026',
        '?action=diag&apiKey=TU_KEY'
      ]
    }, 200);
  }

  const apiKey = String(p.apiKey || '');
  const payload = Object.assign({}, p);
  delete payload.action;
  delete payload.apiKey;

  try {
    const result = handleRequest_({ apiKey, action, payload });
    return jsonOut_(result, 200);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err), stack: (err && err.stack) ? String(err.stack) : null }, 500);
  }
}

// ======== Router ========
function handleRequest_(req) {
  const apiKey = (req && req.apiKey) ? String(req.apiKey) : '';
  const action = (req && req.action) ? String(req.action) : '';
  const payload = (req && req.payload) ? req.payload : {};

  assertAuthorized_(apiKey);

  switch (action) {
    case 'ping':
      return { ok: true, now: new Date().toISOString() };

    case 'diag': {
      const ciclos = getCycles_();
      const ciclo = payload && payload.ciclo_lectivo ? String(payload.ciclo_lectivo).trim() : (ciclos[0] || '');
      const students = getStudentList_({ ciclo_lectivo: ciclo });
      return { ok: true, ciclos, ciclo_usado: ciclo, students_count: (students||[]).length, sample: (students||[]).slice(0, 5) };
    }

    case 'getCycles':
      return { ok: true, cycles: getCycles_() };

    case 'getCatalog':
      return { ok: true, catalog: getCatalog_() };

    case 'getStudentList':
      return { ok: true, students: getStudentList_(payload) };

    case 'getStudentStatus':
      return { ok: true, data: getStudentStatus_(payload) };

    case 'saveStudentStatus':
      return { ok: true, data: saveStudentStatus_(payload) };

    case 'updateStudentOrientation':
      return { ok: true, data: updateStudentOrientation_(payload) };

    case 'syncCatalogRows':
      return { ok: true, data: syncCatalogRows_(payload) };

    case 'getDivisionRiskSummary':
      return { ok: true, data: getDivisionRiskSummary_(payload) };

    case 'closeCycle':
      return { ok: true, data: closeCycle_(payload) };

    case 'rolloverCycle':
      return { ok: true, data: rolloverCycle_(payload) };

    default:
      return { ok: false, error: 'Acción desconocida: ' + action };
  }
}

// ======== Auth ========
function assertAuthorized_(apiKey) {
  const props = PropertiesService.getScriptProperties();
  const realKey = props.getProperty(PROP_API_KEY);
  if (!realKey) throw new Error('No hay API Key configurada. Ejecutá setupApiKey() en el editor.');
  if (!apiKey || apiKey !== realKey) {
    const err = new Error('No autorizado: API Key inválida.');
    err.code = 403;
    throw err;
  }
}

// ======== Supabase REST helpers ========
function getProp_(k) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  return v ? String(v).trim() : '';
}
function requireProp_(k, hint) {
  const v = getProp_(k);
  if (!v) throw new Error('Falta configurar ' + k + (hint ? (' — ' + hint) : ''));
  return v;
}
function supaBase_() { return requireProp_(PROP_SUPABASE_URL, 'Ej: https://xxxx.supabase.co').replace(/\/+$/,''); }
function supaKey_() { return requireProp_(PROP_SUPABASE_SERVICE_KEY, 'Supabase → Settings → API → service_role'); }

function supaFetch_(method, path, query, body, extraHeaders) {
  const url = supaBase_() + path + (query ? ('?' + query) : '');
  const key = supaKey_();
  const headers = {
    apikey: key,
    Authorization: 'Bearer ' + key,
    'Content-Type': 'application/json',
    Accept: 'application/json'
  };
  if (extraHeaders) Object.keys(extraHeaders).forEach(k => headers[k] = extraHeaders[k]);

  const params = { method, muteHttpExceptions: true, headers };
  if (body !== undefined && body !== null) params.payload = JSON.stringify(body);

  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode();
  const text = res.getContentText() || '';
  if (code < 200 || code >= 300) throw new Error('Supabase error ' + code + ': ' + text.slice(0, 500));
  if (!text) return null;
  try { return JSON.parse(text); } catch(e) { return text; }
}

function chunk_(arr, size) {
  const out = [];
  for (let i=0;i<arr.length;i+=size) out.push(arr.slice(i,i+size));
  return out;
}

function getValues_(tableName, filtersQuery) {
  const cols = TABLE_COLS[tableName];
  if (!cols) throw new Error('Tabla no soportada: ' + tableName);
  const qSelect = 'select=' + encodeURIComponent(cols.join(','));
  const q = filtersQuery ? (qSelect + '&' + filtersQuery) : qSelect;
  const data = supaFetch_('GET', '/rest/v1/' + tableName, q, null, null) || [];
  const rows = data.map(o => cols.map(c => (o[c] === undefined ? '' : o[c])));
  return { headers: cols.slice(), rows };
}

function upsertValues_(tableName, headers, rows, onConflict) {
  if (!rows || !rows.length) return { upserted: 0 };
  const cols = headers || TABLE_COLS[tableName];
  const objects = rows.map(r => {
    const o = {};
    cols.forEach((c,i) => { o[c] = (r[i] === '' ? null : r[i]); });
    return o;
  });
  const batches = chunk_(objects, 500);
  batches.forEach(batch => {
    const q = onConflict ? ('on_conflict=' + encodeURIComponent(onConflict)) : '';
    supaFetch_('POST','/rest/v1/' + tableName, q, batch, { Prefer: 'resolution=merge-duplicates,return=minimal' });
  });
  return { upserted: objects.length };
}

function insertRows_(tableName, headers, rows) {
  if (!rows || !rows.length) return { inserted: 0 };
  const cols = headers || TABLE_COLS[tableName];
  const objects = rows.map(r => {
    const o = {};
    cols.forEach((c,i) => { o[c] = (r[i] === '' ? null : r[i]); });
    return o;
  });
  const batches = chunk_(objects, 500);
  batches.forEach(batch => {
    supaFetch_('POST','/rest/v1/' + tableName, '', batch, { Prefer: 'return=minimal' });
  });
  return { inserted: objects.length };
}

function rpc_(name, payload) {
  return supaFetch_('POST', '/rest/v1/rpc/' + name, '', payload || {}, null) || [];
}

// ======== Helpers ========
function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { map[h] = i; });
  return map;
}
function rowToObj_(headers, row) {
  const o = {};
  headers.forEach((h, i) => { o[h] = row[i]; });
  return o;
}
function toBool_(v) {
  if (v === true || v === false) return v;
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  if (s === 'true' || s === 'verdadero' || s === 'si' || s === 'sí' || s === '1' || s === 'x') return true;
  if (s === 'false' || s === 'falso' || s === 'no' || s === '0' || s === '') return false;
  return false;
}
function isoNow_() { return new Date().toISOString(); }
function parseYear_(v) {
  if (v === null || v === undefined) return NaN;
  const s = String(v).trim();
  if (!s) return NaN;
  const m = s.match(/\d+/);
  if (!m) return NaN;
  const n = Number(m[0]);
  return isNaN(n) ? NaN : n;
}
function normalizeOrient_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
function catalogAplicaAStudent_(catMateria, studentGrade, studentOrient) {
  const matOrient = normalizeOrient_(catMateria && catMateria.orientacion);
  if (!matOrient) return true;
  const g = Number(studentGrade || '');
  if (isNaN(g) || g < 4) return false;
  const so = normalizeOrient_(studentOrient);
  if (!so) return false;
  return matOrient === so;
}
function filterCatalogForStudent_(catalog, student) {
  const grade = Number(student && student.anio_actual || '');
  const orient = student ? student.orientacion : '';
  return (catalog || []).filter(m => catalogAplicaAStudent_(m, grade, orient));
}

// Helpers para promo de división (ej: 4°A -> 5°A)
function promoDivision_(division) {
  const s = String(division || '').trim();
  if (!s) return { ok: false, value: s };
  const m = s.match(/^\s*(\d+)\s*(.*)$/);
  if (!m) return { ok: false, value: s };
  const n = Number(m[1]);
  if (isNaN(n)) return { ok: false, value: s };
  const rest = (m[2] || '').trim();
  const next = n + 1;
  const hasDegree = /°/.test(s);
  const deg = hasDegree ? '°' : '';
  let cleanedRest = rest;
  if (cleanedRest.startsWith('°')) cleanedRest = cleanedRest.slice(1).trim();
  return { ok: true, value: `${next}${deg}${cleanedRest ? cleanedRest : ''}`.replace(/\s+/g,' ').trim() };
}

// ======== Actions ========

// Usa RPC get_cycles si existe; si no, fallback con query
function getCycles_() {
  try {
    const rows = rpc_('get_cycles', {});
    const cycles = rows.map(r => String(r.ciclo_lectivo || '').trim()).filter(Boolean);
    const set = {};
    cycles.forEach(c => set[c]=true);
    return Object.keys(set);
  } catch (e) {
    const tmp = getValues_(SHEETS.ESTADO, "select=ciclo_lectivo"); // might fail due to our getValues signature, ignore
    const set = {};
    tmp.rows.forEach(r => { const c = String(r[0]||'').trim(); if (c) set[c]=true; });
    return Object.keys(set);
  }
}

function getCatalog_() {
  const { headers, rows } = getValues_(SHEETS.CATALOGO);
  const idx = headerMap_(headers);
  return rows
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => ({
      id_materia: String(r[idx['id_materia']] || '').trim(),
      nombre: String(r[idx['nombre']] || '').trim(),
      anio: parseYear_(r[idx['anio']]),
      es_troncal: toBool_(r[idx['es_troncal']]),
      orientacion: (idx['orientacion'] !== undefined) ? String(r[idx['orientacion']] || '').trim() : '',
      egresado: (idx['egresado'] !== undefined) ? toBool_(r[idx['egresado']]) : false,
      anio_egreso: (idx['anio_egreso'] !== undefined) ? String(r[idx['anio_egreso']] || '').trim() : ''
    }))
    .filter(m => m.id_materia);
}

function getStudentList_(payload) {
  payload = payload || {};
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const umbral = (payload.umbral !== undefined) ? Number(payload.umbral) : 5;
  if (isNaN(umbral) || umbral < 0) throw new Error('umbral inválido');

  // Leer estudiantes
  const tmp = getValues_(SHEETS.ESTUDIANTES);
  const idx = headerMap_(tmp.headers);

  const students = tmp.rows
    .filter(r => r.some(c => String(c).trim() !== ''))
    .map(r => ({
      id_estudiante: String(r[idx['id_estudiante']] || '').trim(),
      apellido: String(r[idx['apellido']] || '').trim(),
      nombre: String(r[idx['nombre']] || '').trim(),
      anio_actual: Number(r[idx['anio_actual']] || ''),
      division: String(r[idx['division']] || '').trim(),
      turno: String(r[idx['turno']] || '').trim(),
      activo: (idx['activo'] !== undefined) ? toBool_(r[idx['activo']]) : true,
      observaciones: (idx['observaciones'] !== undefined) ? String(r[idx['observaciones']] || '').trim() : '',
      orientacion: (idx['orientacion'] !== undefined) ? String(r[idx['orientacion']] || '').trim() : ''
    }))
    .filter(s => s.id_estudiante)
    .filter(s => s.activo !== false)
    .filter(s => (idx['egresado'] !== undefined) ? (toBool_(String(tmp.rows.find(rr => String(rr[idx['id_estudiante']]||'').trim()===s.id_estudiante)[idx['egresado']])) === false) : true);

  if (!ciclo) return students;

  // Catalog para filtrar por orientación
  const byStudent = {};
  students.forEach(s => { byStudent[s.id_estudiante] = s; });

  const catalogFull = getCatalog_();
  const catalogMap = {};
  catalogFull.forEach(m => { catalogMap[m.id_materia] = m; });

  // Leer estado SOLO del ciclo (no toda la tabla)
  const est = getValues_(SHEETS.ESTADO, 'ciclo_lectivo=eq.' + encodeURIComponent(ciclo));
  const eidx = headerMap_(est.headers);

  const need = {};
  const done = {};
  const needsReview = {};
  const adeudaCount = {};

  est.rows.forEach(r => {
    const sid = String(r[eidx['id_estudiante']] || '').trim();
    if (!sid) return;

    const mid = String(r[eidx['id_materia']] || '').trim();
    if (!mid) return;

    const st = byStudent[sid];
    const cat = catalogMap[mid];
    if (st) {
      if (cat && !catalogAplicaAStudent_(cat, st.anio_actual, st.orientacion)) return;
    }

    const sit = String(r[eidx['situacion_actual']] || '').trim();
    const cond = String(r[eidx['condicion_academica']] || '').trim().toLowerCase();
    const res = (eidx['resultado_cierre'] !== undefined) ? String(r[eidx['resultado_cierre']] || '').trim() : '';

    const resLc = String(res || '').trim().toLowerCase();
    const isAdeuda = (cond === 'adeuda') || (resLc === 'no_aprobada' || resLc === 'no aprobada' || resLc === 'no_aprobo' || resLc === 'no');
    if (isAdeuda) {
      const matYear = cat ? Number(cat.anio || '') : NaN;
      const stYear = st ? Number(st.anio_actual || '') : NaN;
      const hasYears = (!isNaN(matYear) && !isNaN(stYear));
      const countsAsAdeuda = hasYears ? (matYear < stYear) : (sit !== 'proximos_anos' && sit !== 'cursa_primera_vez');
      if (countsAsAdeuda) adeudaCount[sid] = (adeudaCount[sid] || 0) + 1;
    }

    if (sit === 'cursa_primera_vez' || sit === 'recursa' || sit === 'intensifica') {
      need[sid] = (need[sid] || 0) + 1;
      if (res === 'aprobada' || res === 'no_aprobada') done[sid] = (done[sid] || 0) + 1;
    }

    if (sit === 'no_cursa_por_tope') {
      const nunca = (eidx['nunca_cursada'] !== undefined) ? toBool_(r[eidx['nunca_cursada']]) : false;
      if (nunca) needsReview[sid] = true;
    }
  });

  return students.map(s => {
    const total = need[s.id_estudiante] || 0;
    const cerradas = done[s.id_estudiante] || 0;
    const cierreCompleto = (total > 0 && cerradas >= total);
    return Object.assign({}, s, {
      cierre_pendiente: Math.max(0, total - cerradas),
      cierre_completo: cierreCompleto,
      needs_review: !!needsReview[s.id_estudiante],
      adeuda_count: adeudaCount[s.id_estudiante] || 0,
      en_riesgo: (adeudaCount[s.id_estudiante] || 0) >= umbral
    });
  });
}

function getStudentStatus_(payload) {
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const idEst = String(payload.id_estudiante || '').trim();
  if (!ciclo) throw new Error('Falta payload.ciclo_lectivo');
  if (!idEst) throw new Error('Falta payload.id_estudiante');

  // Estudiante
  const students = getStudentList_({});
  const student = students.find(s => s.id_estudiante === idEst) || { id_estudiante: idEst, anio_actual: null, orientacion: '' };

  const catalogFull = getCatalog_();
  const catalog = filterCatalogForStudent_(catalogFull, student);

  const catalogMap = {};
  const allowed = {};
  catalog.forEach(m => { catalogMap[m.id_materia]=m; allowed[m.id_materia]=true; });

  const tmp = getValues_(SHEETS.ESTADO,
    'ciclo_lectivo=eq.' + encodeURIComponent(ciclo) + '&id_estudiante=eq.' + encodeURIComponent(idEst)
  );
  const idx = headerMap_(tmp.headers);

  const materias = tmp.rows
    .map(r => rowToObj_(tmp.headers, r))
    .filter(o => {
      const mid = String(o['id_materia'] || '').trim();
      return !!allowed[mid];
    })
    .map(o => {
      const idMat = String(o['id_materia'] || '').trim();
      const cat = catalogMap[idMat] || {};
      return {
        id_materia: idMat,
        nombre: cat.nombre || '',
        anio: cat.anio || Number(o['anio'] || ''),
        es_troncal: (cat.es_troncal !== undefined) ? cat.es_troncal : toBool_(o['es_troncal']),
        orientacion: cat.orientacion || '',
        condicion_academica: String(o['condicion_academica'] || '').trim(),
        nunca_cursada: toBool_(o['nunca_cursada']),
        situacion_actual: String(o['situacion_actual'] || '').trim(),
        motivo_no_cursa: String(o['motivo_no_cursa'] || '').trim(),
        fecha_actualizacion: o['fecha_actualizacion'] ? new Date(o['fecha_actualizacion']).toISOString() : '',
        usuario: String(o['usuario'] || '').trim(),
        resultado_cierre: (idx['resultado_cierre'] !== undefined) ? String(o['resultado_cierre'] || '').trim() : '',
        ciclo_cerrado: (idx['ciclo_cerrado'] !== undefined) ? toBool_(o['ciclo_cerrado']) : false
      };
    });

  return { ciclo_lectivo: ciclo, estudiante: student, materias };
}

function saveStudentStatus_(payload) {
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const idEst = String(payload.id_estudiante || '').trim();
  const usuario = String(payload.usuario || 'web').trim();
  const updates = payload.updates || [];
  if (!ciclo) throw new Error('Falta payload.ciclo_lectivo');
  if (!idEst) throw new Error('Falta payload.id_estudiante');
  if (!Array.isArray(updates) || updates.length === 0) throw new Error('Falta payload.updates (array)');

  // Traer actuales (solo ciclo+estudiante)
  const tmp = getValues_(SHEETS.ESTADO,
    'ciclo_lectivo=eq.' + encodeURIComponent(ciclo) + '&id_estudiante=eq.' + encodeURIComponent(idEst)
  );
  const headers = tmp.headers;
  const rows = tmp.rows;
  const idx = headerMap_(headers);

  const existingByMid = {};
  rows.forEach(r => {
    const mid = String(r[idx['id_materia']] || '').trim();
    if (mid) existingByMid[mid] = r;
  });

  const now = new Date().toISOString();
  const upsertRows = [];
  const auditRows = [];

  updates.forEach(u => {
    const idMat = String(u.id_materia || '').trim();
    if (!idMat) return;
    const fields = u.fields || {};

    const prev = existingByMid[idMat] ? existingByMid[idMat].slice() : null;
    const newRow = prev ? prev.slice() : headers.map(() => '');

    newRow[idx['ciclo_lectivo']] = ciclo;
    newRow[idx['id_estudiante']] = idEst;
    newRow[idx['id_materia']] = idMat;

    Object.keys(fields).forEach(k => {
      if (idx[k] === undefined) return;
      let v = fields[k];
      if (k === 'nunca_cursada' || k === 'ciclo_cerrado') v = !!v;
      newRow[idx[k]] = v;
    });

    if (idx['fecha_actualizacion'] !== undefined) newRow[idx['fecha_actualizacion']] = now;
    if (idx['usuario'] !== undefined) newRow[idx['usuario']] = usuario;

    // Auditoría
    if (prev) {
      Object.keys(fields).forEach(k => {
        if (idx[k] === undefined) return;
        const beforeVal = prev[idx[k]];
        const afterVal = fields[k];
        if (String(beforeVal) !== String(afterVal)) {
          auditRows.push([now, ciclo, idEst, idMat, k, String(beforeVal||''), String(afterVal||''), usuario]);
        }
      });
    } else {
      Object.keys(fields).forEach(k => {
        if (idx[k] === undefined) return;
        auditRows.push([now, ciclo, idEst, idMat, k, '', String(fields[k]||''), usuario]);
      });
    }

    upsertRows.push(newRow);
  });

  upsertValues_(SHEETS.ESTADO, headers, upsertRows, 'ciclo_lectivo,id_estudiante,id_materia');
  if (auditRows.length) insertRows_(SHEETS.AUDITORIA, TABLE_COLS['auditoria'], auditRows);

  return getStudentStatus_({ ciclo_lectivo: ciclo, id_estudiante: idEst });
}

function updateStudentOrientation_(payload) {
  const idEst = String(payload.id_estudiante || '').trim();
  const orient = String(payload.orientacion || '').trim();
  const usuario = String(payload.usuario || 'web').trim();
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  if (!idEst) throw new Error('Falta payload.id_estudiante');

  const tmp = getValues_(SHEETS.ESTUDIANTES, 'id_estudiante=eq.' + encodeURIComponent(idEst));
  const headers = tmp.headers;
  const rows = tmp.rows;
  const idx = headerMap_(headers);
  if (!rows.length) throw new Error('No se encontró el estudiante: ' + idEst);

  const row = rows[0].slice();
  const before = (idx['orientacion'] !== undefined) ? String(row[idx['orientacion']] || '').trim() : '';
  row[idx['orientacion']] = orient;

  upsertValues_(SHEETS.ESTUDIANTES, headers, [row], 'id_estudiante');

  const now = new Date().toISOString();
  insertRows_(SHEETS.AUDITORIA, TABLE_COLS['auditoria'], [[now, ciclo || '', idEst, '', 'orientacion', before, orient, usuario]]);

  return { id_estudiante: idEst, orientacion: orient };
}

function syncCatalogRows_(payload) {
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const idEst = String(payload.id_estudiante || '').trim();
  const usuario = String(payload.usuario || 'web').trim();
  if (!ciclo) throw new Error('Falta payload.ciclo_lectivo');
  if (!idEst) throw new Error('Falta payload.id_estudiante');

  const students = getStudentList_({});
  const student = students.find(s => s.id_estudiante === idEst) || { id_estudiante: idEst, anio_actual: null, orientacion: '' };
  const grade = Number(student.anio_actual || '');

  const catalogFull = getCatalog_();
  const allowedCatalog = filterCatalogForStudent_(catalogFull, student);

  const yearCatalog = (!isNaN(grade) && grade > 0)
    ? allowedCatalog.filter(m => Number(m.anio || '') === grade)
    : [];

  const tmp = getValues_(SHEETS.ESTADO,
    'ciclo_lectivo=eq.' + encodeURIComponent(ciclo) + '&id_estudiante=eq.' + encodeURIComponent(idEst)
  );
  const headers = tmp.headers;
  const rows = tmp.rows;
  const idx = headerMap_(headers);

  const existing = {};
  rows.forEach(r => {
    const mid = String(r[idx['id_materia']] || '').trim();
    if (mid) existing[mid]=true;
  });

  // Crear faltantes (solo año actual)
  const now = new Date().toISOString();
  const newRows = [];
  let added = 0;

  yearCatalog.forEach(m => {
    const mid = String(m.id_materia||'').trim();
    if (!mid) return;
    if (existing[mid]) return;

    const obj = {};
    headers.forEach(h => obj[h]='');
    obj['ciclo_lectivo']=ciclo;
    obj['id_estudiante']=idEst;
    obj['id_materia']=mid;
    obj['condicion_academica']='adeuda';
    obj['nunca_cursada']=true;
    obj['situacion_actual']='cursa_primera_vez';
    obj['motivo_no_cursa']='';
    obj['resultado_cierre']='';
    obj['ciclo_cerrado']=false;
    obj['fecha_actualizacion']=now;
    obj['usuario']=usuario;

    newRows.push(headers.map(h => obj[h]));
    added++;
  });

  if (newRows.length) upsertValues_(SHEETS.ESTADO, headers, newRows, 'ciclo_lectivo,id_estudiante,id_materia');
  return { added, status: getStudentStatus_({ ciclo_lectivo: ciclo, id_estudiante: idEst }) };
}

function getDivisionRiskSummary_(payload) {
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const umbral = (payload.umbral !== undefined) ? Number(payload.umbral) : 5;
  if (!ciclo) throw new Error('Falta payload.ciclo_lectivo');
  if (isNaN(umbral) || umbral < 0) throw new Error('umbral inválido');

  const students = getStudentList_({});
  const byId = {};
  students.forEach(s => { byId[s.id_estudiante]=s; });

  const catalogFull = getCatalog_();
  const catalogMap = {};
  catalogFull.forEach(m => { catalogMap[m.id_materia]=m; });

  const est = getValues_(SHEETS.ESTADO, 'ciclo_lectivo=eq.' + encodeURIComponent(ciclo));
  const idx = headerMap_(est.headers);

  const adeudaCount = {};
  const hasAny = {};

  est.rows.forEach(r => {
    const sid = String(r[idx['id_estudiante']] || '').trim();
    const st = byId[sid];
    if (!st) return;

    const mid = String(r[idx['id_materia']] || '').trim();
    if (!mid) return;
    const cat = catalogMap[mid];
    if (cat && !catalogAplicaAStudent_(cat, st.anio_actual, st.orientacion)) return;

    hasAny[sid] = true;
    const cond = String(r[idx['condicion_academica']] || '').trim().toLowerCase();
    const sit = (idx['situacion_actual'] !== undefined) ? String(r[idx['situacion_actual']] || '').trim() : '';

    const matYear = cat ? Number(cat.anio || '') : NaN;
    const stYear = st ? Number(st.anio_actual || '') : NaN;
    const futureByYear = (!isNaN(matYear) && !isNaN(stYear) && matYear > stYear);

    if (cond === 'adeuda' && sit !== 'proximos_anos' && !futureByYear) {
      adeudaCount[sid] = (adeudaCount[sid] || 0) + 1;
    }
  });

  const groups = {};
  students.forEach(s => {
    const key = `${s.division || '—'}|${s.turno || ''}`;
    if (!groups[key]) groups[key] = { division: s.division || '—', turno: s.turno || '', total_estudiantes: 0, en_riesgo: 0, sin_datos: 0 };
    groups[key].total_estudiantes++;
    const cnt = adeudaCount[s.id_estudiante] || 0;
    if (cnt >= umbral) groups[key].en_riesgo++;
    if (!hasAny[s.id_estudiante]) groups[key].sin_datos++;
  });

  const result = Object.values(groups).sort((a,b) => String(a.division).localeCompare(String(b.division)) || String(a.turno).localeCompare(String(b.turno)));
  return { ciclo_lectivo: ciclo, umbral, divisiones: result };
}

function closeCycle_(payload) {
  const ciclo = String(payload.ciclo_lectivo || '').trim();
  const idEst = payload.id_estudiante ? String(payload.id_estudiante).trim() : '';
  const usuario = String(payload.usuario || 'cierre').trim();
  const marcarCerrado = (payload.marcar_cerrado !== undefined) ? toBool_(payload.marcar_cerrado) : true;
  if (!ciclo) throw new Error('Falta payload.ciclo_lectivo');

  let fq = 'ciclo_lectivo=eq.' + encodeURIComponent(ciclo);
  if (idEst) fq += '&id_estudiante=eq.' + encodeURIComponent(idEst);

  const tmp = getValues_(SHEETS.ESTADO, fq);
  const headers = tmp.headers;
  const rows = tmp.rows;
  const idx = headerMap_(headers);

  const now = new Date().toISOString();
  let updated = 0;
  let scanned = 0;

  const outRows = [];

  rows.forEach(row0 => {
    const row = row0.slice();
    scanned++;

    const rc = String(row[idx['resultado_cierre']] || '').trim().toLowerCase();
    if (!rc) return;

    const aprobo = (rc === 'aprobada' || rc === 'aprobo' || rc === 'aprobó' || rc === 'si' || rc === 'sí');
    const noAprobo = (rc === 'no_aprobada' || rc === 'no aprobada' || rc === 'no_aprobo' || rc === 'no aprobó' || rc === 'no');
    if (!(aprobo || noAprobo)) return;

    if (aprobo) {
      row[idx['condicion_academica']] = 'aprobada';
      if (idx['situacion_actual'] !== undefined) row[idx['situacion_actual']] = '';
      if (idx['motivo_no_cursa'] !== undefined) row[idx['motivo_no_cursa']] = '';
      if (idx['nunca_cursada'] !== undefined) row[idx['nunca_cursada']] = false;
    } else {
      row[idx['condicion_academica']] = 'adeuda';
      if (idx['nunca_cursada'] !== undefined) row[idx['nunca_cursada']] = false;
    }

    if (marcarCerrado && idx['ciclo_cerrado'] !== undefined) row[idx['ciclo_cerrado']] = true;
    if (idx['fecha_actualizacion'] !== undefined) row[idx['fecha_actualizacion']] = now;
    if (idx['usuario'] !== undefined) row[idx['usuario']] = usuario;

    updated++;
    outRows.push(row);
  });

  if (outRows.length) upsertValues_(SHEETS.ESTADO, headers, outRows, 'ciclo_lectivo,id_estudiante,id_materia');
  const status = idEst ? getStudentStatus_({ ciclo_lectivo: ciclo, id_estudiante: idEst }) : null;
  return { ciclo_lectivo: ciclo, id_estudiante: idEst || null, filas_revisadas: scanned, filas_actualizadas: updated, status };
}

// Rollover simplificado (crea ciclo destino SIN tocar ciclos anteriores)
// payload: {ciclo_origen, ciclo_destino, usuario}
function rolloverCycle_(payload) {
  const origen = String(payload.ciclo_origen || '').trim();
  const destino = String(payload.ciclo_destino || '').trim();
  const usuario = String(payload.usuario || 'rollover').trim();
  if (!origen) throw new Error('Falta payload.ciclo_origen');
  if (!destino) throw new Error('Falta payload.ciclo_destino');
  if (origen === destino) throw new Error('ciclo_origen y ciclo_destino no pueden ser iguales');

  // Check pendientes
  const studentsWithFlags = getStudentList_({ ciclo_lectivo: origen });
  const pendientes = studentsWithFlags.filter(s => Number(s.cierre_pendiente || 0) > 0);
  if (pendientes.length > 0) {
    throw new Error('No se puede crear el ciclo nuevo: hay estudiantes con materias sin cierre en el ciclo ' + origen + '.');
  }

  const students = getStudentList_({});
  const catalog = getCatalog_();

  // Leer estado histórico (solo para construir mapas) — con 10k es ok
  const allEstado = getValues_(SHEETS.ESTADO);
  const headers = allEstado.headers;
  const rows = allEstado.rows;
  const idx = headerMap_(headers);

  const approvedMap = {}; // sid|mid
  const regularMap = {};  // sid|mid
  const existsDest = {};  // sid|mid

  rows.forEach(r => {
    const c = String(r[idx['ciclo_lectivo']] || '').trim();
    const sid = String(r[idx['id_estudiante']] || '').trim();
    const mid = String(r[idx['id_materia']] || '').trim();
    if (!c || !sid || !mid) return;
    const key = sid + '|' + mid;

    if (c === destino) { existsDest[key] = true; return; }

    const cond = String(r[idx['condicion_academica']] || '').trim().toLowerCase();
    const sit = String(r[idx['situacion_actual']] || '').trim();
    const resCierre = (idx['resultado_cierre'] !== undefined) ? String(r[idx['resultado_cierre']] || '').trim().toLowerCase() : '';

    if (cond === 'aprobada' || resCierre === 'aprobada' || resCierre === 'aprobo' || resCierre === 'aprobó') approvedMap[key] = true;
    if (sit === 'cursa_primera_vez' || sit === 'recursa') regularMap[key] = true;
  });

  const now = new Date().toISOString();
  const newRows = [];
  let created = 0;

  students.forEach(s => {
    const sid = s.id_estudiante;
    const allowedCatalog = filterCatalogForStudent_(catalog, s);
    allowedCatalog.forEach(m => {
      const mid = String(m.id_materia||'').trim();
      if (!mid) return;
      const key = sid + '|' + mid;
      if (existsDest[key]) return;
      if (approvedMap[key]) return;

      const everRegular = !!regularMap[key];

      const obj = {};
      headers.forEach(h => obj[h]='');
      obj['ciclo_lectivo']=destino;
      obj['id_estudiante']=sid;
      obj['id_materia']=mid;
      obj['condicion_academica']='adeuda';
      obj['nunca_cursada']=!everRegular;
      obj['situacion_actual']='no_cursa_otro_motivo';
      obj['motivo_no_cursa']='';
      obj['resultado_cierre']='';
      obj['ciclo_cerrado']=false;
      obj['fecha_actualizacion']=now;
      obj['usuario']=usuario;

      newRows.push(headers.map(h => obj[h]));
      created++;
    });
  });

  if (newRows.length) upsertValues_(SHEETS.ESTADO, headers, newRows, 'ciclo_lectivo,id_estudiante,id_materia');
  return { ciclo_origen: origen, ciclo_destino: destino, filas_creadas: created };
}

// ======== Output ========
function jsonOut_(obj, statusCode) {
  const payload = Object.assign({ http_status: statusCode }, obj);
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
