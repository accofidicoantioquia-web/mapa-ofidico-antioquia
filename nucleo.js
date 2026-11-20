/* Núcleo compartido del mapa vivo: clasificación de corredores, reconocimiento de veredas y generación de datos.
   Lo usan index.html (visor) y actualizador.html (actualizador). No contiene datos de personas. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Nucleo = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // Tipo de entidad que se escribe en la base para casos de zona urbana (sin polígono: se marcan con un pin)
  var ENT_URBANA = 'ZONA URBANA (sin polígono; se marca con pin)';
  var MIN_CASOS_PIN = 2;   // una zona urbana se marca con pin cuando acumula este número de casos (criterio de vereda individual)

  // ---------------------------------------------------------------- utilidades de texto
  function norm(s) {
    s = String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
    return s.replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
  }
  var PFX = /^(VEREDA|VDA|V\.?|CORREGIMIENTO|CGTO|CTO|CRG)\s+/;
  var JUNK = { 'DESCONOCIDA': 1, 'RURAL': 1, 'NO APLICA': 1, 'NA': 1, 'N A': 1, 'NR': 1, 'N R': 1, 'NO REPORTA': 1,
               'SIN DATO': 1, 'SIN DATOS': 1, 'NINGUNA': 1, 'NINGUNO': 1, 'NO SABE': 1, 'SD': 1, 'S D': 1, 'NO INFORMA': 1 };
  // Devuelve {texto, motivoInvalido}. texto = nombre reportado limpio (mayúsculas, sin prefijo VEREDA/VDA/…)
  function limpiarVereda(raw) {
    if (raw == null) return { texto: '', invalido: 'VACIO' };
    var s = String(raw).trim().replace(/\s+/g, ' ').toUpperCase();
    if (!s) return { texto: '', invalido: 'VACIO' };
    s = s.replace(PFX, '');
    var k = norm(s);
    if (!k) return { texto: s, invalido: 'INVALIDO' };
    if (JUNK[k] || /^[0-9 ]+$/.test(k) || /^X+$/.test(k)) return { texto: s, invalido: 'INVALIDO' };
    return { texto: s, invalido: null };
  }

  // ---------------------------------------------------------------- fechas (fecha de la última información del corte)
  var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  function isoFecha(y, m, d) {
    y = +y; m = +m; d = +d;
    if (!(y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1)) return null;
    var dm = [31, (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0)) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1];
    if (d > dm) return null;
    return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2);
  }
  // 'dmy' (día/mes/año, lo habitual en Colombia) o 'mdy'; se deduce de los valores: un primer número > 12 es día; un segundo número > 12 es día.
  function ordenFechas(textos) {
    var a = 0, b = 0;
    textos.forEach(function (s) { var m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.]\d{4}/.exec(s); if (m) { if (+m[1] > 12) a++; if (+m[2] > 12) b++; } });
    return b > a ? 'mdy' : 'dmy';
  }
  // v: Date (celda de fecha de Excel), número de serie de Excel o texto (AAAA-MM-DD, DD/MM/AAAA, con o sin hora). Devuelve 'AAAA-MM-DD' o null.
  function parseFecha(v, orden) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : isoFecha(v.getUTCFullYear(), v.getUTCMonth() + 1, v.getUTCDate());
    if (typeof v === 'number') {
      if (!(v >= 36526 && v <= 73050)) return null;               // serie de Excel entre 2000 y 2099
      var dt = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
      return isoFecha(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate());
    }
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/.exec(s))) return isoFecha(m[1], m[2], m[3]);
    if ((m = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:[T\s].*)?$/.exec(s))) return orden === 'mdy' ? isoFecha(m[3], m[1], m[2]) : isoFecha(m[3], m[2], m[1]);
    if (/^\d{5}$/.test(s)) return parseFecha(parseInt(s, 10), orden);
    return null;
  }
  // La fecha más reciente de una lista de valores, sin contar fechas posteriores a `tope` ('AAAA-MM-DD', normalmente hoy) ni anteriores al 2000.
  function fechaMaxima(valores, tope) {
    var textos = {}, orden, mejor = null, cache = {};
    valores.forEach(function (v) { if (typeof v === 'string' && v.trim()) textos[v.trim()] = 1; });
    orden = ordenFechas(Object.keys(textos));
    valores.forEach(function (v) {
      var f;
      if (typeof v === 'string') { var s = v.trim(); if (!s) return; f = Object.prototype.hasOwnProperty.call(cache, s) ? cache[s] : (cache[s] = parseFecha(s, orden)); }
      else f = parseFecha(v, orden);
      if (!f || f < '2000-01-01' || (tope && f > tope)) return;
      if (!mejor || f > mejor) mejor = f;
    });
    return mejor;
  }
  function fechaLarga(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? (+m[3]) + ' de ' + MESES[+m[2] - 1] + ' de ' + m[1] : '';
  }

  // ---------------------------------------------------------------- huella de caso (sin datos personales)
  function cyrb53(str, seed) {
    seed = seed || 0;
    var h1 = 0xdeadbeef ^ seed, h2 = 0x41c6ce57 ^ seed;
    for (var i = 0, ch; i < str.length; i++) {
      ch = str.charCodeAt(i);
      h1 = Math.imul(h1 ^ ch, 2654435761);
      h2 = Math.imul(h2 ^ ch, 1597334677);
    }
    h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507);
    h1 ^= Math.imul(h2 ^ (h2 >>> 13), 3266489909);
    h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507);
    h2 ^= Math.imul(h1 ^ (h1 >>> 13), 3266489909);
    return 4294967296 * (2097151 & h2) + (h1 >>> 0);
  }
  function normId(v) {
    if (v == null) return '';
    var s = String(v).trim().toUpperCase();
    if (/^\d+(\.0+)?$/.test(s)) return String(parseInt(s, 10));
    return s;
  }
  // Campos estables entre cortes (no incluye fecha de notificación, nombre común ni vereda, que SIVIGILA corrige con frecuencia)
  function huellaBase(c) {
    var txt = [c.ano, c.semana, c.cod_mun_o, c.edad_, c.uni_med_, c.sexo_, c.age_agrg].map(normId).join('|');
    return cyrb53(txt).toString(16).padStart(14, '0');
  }

  // ---------------------------------------------------------------- lectura de CSV
  // Lee CSV. Si se entrega onRow(fila, numero), las filas no se acumulan (se procesan y se descartan una a una).
  function parseCSV(text, onRow, onHeader) {
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    var nl = text.indexOf('\n'); var first = nl < 0 ? text : text.slice(0, nl);
    var inQ = false, cnt = { ',': 0, ';': 0, '\t': 0, '|': 0 };
    for (var i = 0; i < first.length; i++) { var ch = first[i]; if (ch === '"') inQ = !inQ; else if (!inQ && cnt[ch] !== undefined) cnt[ch]++; }
    var delim = ','; var best = -1; ['\t', ';', '|', ','].forEach(function (d) { if (cnt[d] > best) { best = cnt[d]; delim = d; } });
    var rows = [], headers = null, n = 0, row = [], field = '', q = false;
    function fin() {
      row.push(field); field = '';
      if (row.length > 1 || row[0] !== '') {
        if (!headers) { headers = row; if (onHeader) onHeader(headers); }
        else { n++; if (onRow) onRow(row, n); else rows.push(row); }
      }
      row = [];
    }
    for (var p = 0; p < text.length; p++) {
      var c = text[p];
      if (q) {
        if (c === '"') { if (text[p + 1] === '"') { field += '"'; p++; } else q = false; } else field += c;
      } else if (c === '"') q = true;
      else if (c === delim) { row.push(field); field = ''; }
      else if (c === '\n') fin();
      else if (c === '\r') { /* ignorar */ }
      else field += c;
    }
    if (field.length || row.length) fin();
    return { delim: delim, headers: headers || [], rows: rows, filas: n };
  }

  // Columnas de interés del archivo SIVIGILA (el resto, incluidos los datos personales, nunca se lee)
  var COLS = {
    cod_eve: ['CODEVE'], ano: ['ANO', 'ANIO', 'AO', 'AAO', 'YEAR'], semana: ['SEMANA'], fec_not: ['FECNOT'],
    cod_dpto_o: ['CODDPTOO'], cod_mun_o: ['CODMUNO'], nmun_proce: ['NMUNPROCE'], vereda_: ['VEREDA'],
    cen_pobla_: ['CENPOBLA'], bar_ver_: ['BARVER'], edad_: ['EDAD'], uni_med_: ['UNIMED'], sexo_: ['SEXO'],
    age_agrg: ['AGEAGRG'], age_agrn: ['AGEAGRN'], ajuste_: ['AJUSTE']
  };
  var COLS_REQUERIDAS = ['ano', 'semana', 'cod_mun_o', 'vereda_', 'edad_', 'uni_med_', 'sexo_', 'age_agrg', 'age_agrn'];
  function mapearColumnas(headers) {
    var idx = {}, hk = headers.map(function (h) { return norm(h).replace(/ /g, ''); });
    Object.keys(COLS).forEach(function (name) {
      for (var j = 0; j < COLS[name].length; j++) {
        var pos = hk.indexOf(COLS[name][j].replace(/ /g, ''));
        if (pos >= 0) { idx[name] = pos; break; }
      }
    });
    var faltan = COLS_REQUERIDAS.filter(function (n) { return idx[n] === undefined; });
    return { idx: idx, faltan: faltan };
  }
  // fila (array) -> caso con solo las columnas de interés
  function extraerCaso(fila, idx) {
    var o = {};
    Object.keys(idx).forEach(function (n) { var v = fila[idx[n]]; o[n] = v == null ? '' : (typeof v === 'object' && v.text ? v.text : v); });
    return o;
  }

  // ---------------------------------------------------------------- índices de referencia
  function construirIndices(REF, POLIGONOS) {
    var OF = {}, DANE = {}, TODOS = {}, KEY2IDX = {}, DANEALL = {};
    Object.keys(REF.OFICIAL).forEach(function (cod) {
      OF[cod] = {};
      REF.OFICIAL[cod].forEach(function (n) { var k = norm(n); if (k && !OF[cod][k]) OF[cod][k] = n; (TODOS[k] = TODOS[k] || {})[cod] = 1; });
    });
    POLIGONOS.forEach(function (p, i) {
      var cod = p[0], k = p[3]; (DANE[cod] = DANE[cod] || {})[k] = p[2]; (DANEALL[k] = DANEALL[k] || {})[cod] = 1; (TODOS[k] = TODOS[k] || {})[cod] = 1;
      if (p[4] === 1) KEY2IDX[cod + '|' + k] = i;
    });
    var alias = {};
    REF.ALIAS.forEach(function (a) { alias[a[0] + '|' + a[1]] = a[2]; });
    // Veredas del listado oficial cuyo polígono DANE está en un municipio vecino (REF.PRESTADOS = [municipio oficial, nombre, municipio del polígono]):
    // se dibujan con ese polígono y conservan su municipio oficial. EFF: índice del polígono → clave oficial; PREST: clave oficial → clave real del polígono.
    var EFF = {}, PREST = {};
    (REF.PRESTADOS || []).forEach(function (e) {
      var oficial = e[0] + '|' + e[1], real = e[2] + '|' + e[1], i = KEY2IDX[real];
      if (i === undefined || KEY2IDX[oficial] !== undefined) return;     // sin polígono de origen, o el municipio oficial ya tiene el suyo
      KEY2IDX[oficial] = i; EFF[i] = oficial; PREST[oficial] = real;
    });
    return { OF: OF, DANE: DANE, DANEALL: DANEALL, TODOS: TODOS, KEY2IDX: KEY2IDX, ALIAS: alias, EFF: EFF, PREST: PREST };
  }

  // ---------------------------------------------------------------- similitud de nombres
  var STOP = { EL: 1, LA: 1, LOS: 1, LAS: 1, DE: 1, DEL: 1, Y: 1, E: 1 };
  function lev(a, b) {
    var m = a.length, n = b.length; if (!m) return n; if (!n) return m;
    var prev = new Array(n + 1), cur = new Array(n + 1), i, j;
    for (j = 0; j <= n; j++) prev[j] = j;
    for (i = 1; i <= m; i++) {
      cur[0] = i;
      for (j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      var t = prev; prev = cur; cur = t;
    }
    return prev[n];
  }
  function tokens(k) { return k.split(' ').filter(function (t) { return t && !STOP[t]; }); }
  // Relación entre el nombre reportado (k) y un nombre de referencia (ck)
  function relacion(k, ck) {
    var tk = tokens(k), tc = tokens(ck);
    if (tk.slice().sort().join(' ') === tc.slice().sort().join(' ') || k.replace(/ /g, '') === ck.replace(/ /g, '')) return { rel: 'VARIANTE', score: 0.99 };
    var d = lev(k, ck), sim = 1 - d / Math.max(k.length, ck.length);
    if (d <= 2 && Math.max(k.length, ck.length) >= 5) return { rel: 'TYPO', score: 0.9 + sim / 20 };
    if (tk.length && tc.length) {
      var sub = function (a, b) { return a.every(function (t) { return b.indexOf(t) >= 0; }); };
      if (sub(tk, tc) || sub(tc, tk)) return { rel: 'PARCIAL', score: 0.8 };
    }
    if (sim >= 0.75) return { rel: 'TYPO', score: sim };
    return null;
  }
  // Equivalencias posibles dentro del municipio (listado oficial + capa DANE)
  function sugerencias(cod, k, IX, max) {
    var out = {}, add = function (ck, nombre, fuente) {
      if (ck === k) return;
      var r = relacion(k, ck); if (!r) return;
      var o = out[ck] || (out[ck] = { nombre: nombre, k: ck, rel: r.rel, score: r.score, oficial: false, poligono: false });
      if (fuente === 'O') o.oficial = true; else o.poligono = true;
    };
    var of = IX.OF[cod] || {}, dn = IX.DANE[cod] || {};
    Object.keys(of).forEach(function (ck) { add(ck, of[ck], 'O'); });
    Object.keys(dn).forEach(function (ck) { add(ck, dn[ck], 'D'); });
    if (dn[k] && !of[k]) out[k] = { nombre: dn[k], k: k, rel: 'EXACTO_DANE', score: 1, oficial: false, poligono: true };
    var lst = Object.keys(out).map(function (x) { return out[x]; }).sort(function (a, b) { return b.score - a.score; });
    return lst.slice(0, max || 6);
  }
  function otrosMunicipios(cod, k, IX, REF) {
    var t = IX.TODOS[k] || {}; return Object.keys(t).filter(function (c) { return c !== cod; }).map(function (c) { return { cod: c, nombre: (REF.MUN[c] || [c])[0] }; });
  }

  // ---------------------------------------------------------------- clasificación (misma lógica del análisis espacial)
  // datos: {casos:{'cod|K':n}, resguardos:{'cod|K':n}, urbanas:{'cod|K':{n,nombre}}}
  function clasificar(datos, POLIGONOS, COLIND, REF, IX) {
    var EXCL = {}; REF.EXCL.forEach(function (e) { EXCL[e[0] + '|' + e[1]] = 1; });
    var DISPLAY = {}; REF.DISPLAY.forEach(function (e) { DISPLAY[e[0] + '|' + e[1]] = e[2]; });
    if (!IX) IX = construirIndices(REF, POLIGONOS);
    var KEY2IDX = IX.KEY2IDX;
    var mostrar = datos.mostrar || {};
    // polígonos prestados activos: solo si el municipio del polígono no tiene casos con ese mismo nombre (si los tiene, el polígono es de ese municipio)
    var efectiva = {};
    Object.keys(IX.EFF || {}).forEach(function (i) { var real = POLIGONOS[i][0] + '|' + POLIGONOS[i][3]; if (!(datos.casos[real] > 0)) efectiva[i] = IX.EFF[i]; });
    var casos = {};
    Object.keys(datos.casos).forEach(function (k) {
      var i = KEY2IDX[k];
      if (EXCL[k] || i === undefined) return;
      if (IX.EFF && IX.EFF[i] === k && efectiva[i] === undefined) return;      // polígono prestado en conflicto: no se dibuja
      casos[k] = datos.casos[k];
    });
    var keyOf = function (i) { return efectiva[i] || (POLIGONOS[i][0] + '|' + POLIGONOS[i][3]); };
    var enCorredor = {}, paresMapa = 0, paresTodos = 0, pares = [];
    COLIND.forEach(function (pr) {
      var a = pr[0], b = pr[1], ka = keyOf(a), kb = keyOf(b);
      if (EXCL[ka] || EXCL[kb]) return;
      if (ka.split('|')[0] === kb.split('|')[0]) return;                       // mismo municipio (oficial): no es frontera
      if (!(datos.casos[ka] > 0 && datos.casos[kb] > 0)) return;
      paresTodos++;                                           // convención del manuscrito: todos los polígonos
      pares.push([a, b]);
      if (POLIGONOS[a][4] === 1 && POLIGONOS[b][4] === 1) { paresMapa++; enCorredor[ka] = 1; enCorredor[kb] = 1; }   // convención del mapa
    });
    var feats = [];
    var claves = Object.keys(casos).sort();
    ['VEREDA_CORREDOR', 'VEREDA_INDIVIDUAL'].forEach(function (tipo) {
      claves.forEach(function (k) {
        var t = enCorredor[k] ? 'VEREDA_CORREDOR' : (casos[k] >= 2 ? 'VEREDA_INDIVIDUAL' : null);
        if (t !== tipo) return;
        var ix = KEY2IDX[k], p = POLIGONOS[ix], cm = p[0], nm = p[1];
        if (efectiva[ix] === k) { cm = k.split('|')[0]; nm = (REF.MUN[cm] || [nm])[0]; }       // polígono prestado: municipio oficial
        feats.push({ type: 'Feature',
          properties: { cod_mpio: cm, nom_mpio: nm, nom_ver: mostrar[k] || DISPLAY[k] || p[2], casos: casos[k], tipo_entidad: tipo, clave: k },
          geometry: { type: p[5] === 'P' ? 'Polygon' : 'MultiPolygon', coordinates: p[6] } });
      });
    });
    REF.RESGUARDOS.forEach(function (r) {
      var n = (datos.resguardos || {})[r.cod + '|' + r.k];
      if (n > 0) feats.push({ type: 'Feature', properties: { cod_mpio: r.cod, nom_mpio: r.mun, nom_ver: r.ver, casos: n, tipo_entidad: 'RESGUARDO', clave: r.cod + '|' + r.k }, geometry: r.geometry });
    });
    feats.forEach(function (f, i) { f.id = String(i); });
    var cuenta = { VEREDA_CORREDOR: 0, VEREDA_INDIVIDUAL: 0, RESGUARDO: 0 }, suma = { VEREDA_CORREDOR: 0, VEREDA_INDIVIDUAL: 0, RESGUARDO: 0 };
    feats.forEach(function (f) { cuenta[f.properties.tipo_entidad]++; suma[f.properties.tipo_entidad] += f.properties.casos; });
    // Zonas urbanas (sin polígono): un pin por municipio, ubicado en su cabecera (DIVIPOLA), con las zonas que acumulan 2 o más casos
    var porMun = {}, pocos = 0;
    Object.keys(datos.urbanas || {}).sort().forEach(function (k) {
      var u = datos.urbanas[k], cod = k.split('|')[0];
      if (!(u.n >= MIN_CASOS_PIN)) { pocos += u.n; return; }
      var cab = (REF.CABECERAS || {})[cod];
      if (!cab) { pocos += u.n; return; }
      var m = porMun[cod] || (porMun[cod] = { cod: cod, nom_mpio: (REF.MUN[cod] || [cod])[0], lat: cab[0], lon: cab[1], zonas: [], casos: 0 });
      m.zonas.push({ nombre: u.nombre, casos: u.n, clave: k }); m.casos += u.n;
    });
    var pines = Object.keys(porMun).sort().map(function (c) { return porMun[c]; });
    cuenta.ZONA_URBANA = pines.length; suma.ZONA_URBANA = pines.reduce(function (a, p) { return a + p.casos; }, 0);
    return { features: feats, paresMapa: paresMapa, paresTodos: paresTodos, pares: pares, cuenta: cuenta, suma: suma, pines: pines, urbanasSinPin: pocos };
  }

  // ---------------------------------------------------------------- datos.js a partir de las filas de la base
  // filas: [{cod:'05045', ajustada:'...', entidad:'...'}]; extra.alias: {'cod|K': 'K_DANE'} (decisiones del usuario)
  function generarDatos(filas, REF, IX, extra) {
    extra = extra || {};
    var EXCL = {}; REF.EXCL.forEach(function (e) { EXCL[e[0] + '|' + e[1]] = 1; });
    var aliasDin = extra.alias || {};
    var resg = {}; REF.RESGUARDOS.forEach(function (r) { resg[r.cod + '|' + r.k] = 1; });
    var casos = {}, resguardos = {}, mostrar = {}, total = 0, urbanas = {}, nomUrb = {};
    filas.forEach(function (f) {
      var k = norm(f.ajustada);
      if (k === 'NO REPORTA') return;
      total++;
      var key0 = f.cod + '|' + k;
      if (f.entidad === ENT_URBANA) {                        // zona urbana: se cuenta aparte (no es una vereda)
        var u = urbanas[key0] || (urbanas[key0] = { n: 0, nombre: '' }), sp = nomUrb[key0] || (nomUrb[key0] = {});
        u.n++; sp[f.ajustada] = (sp[f.ajustada] || 0) + 1;
        return;
      }
      var key = IX.ALIAS[key0] ? f.cod + '|' + IX.ALIAS[key0] : (aliasDin[key0] ? f.cod + '|' + aliasDin[key0] : key0);
      if (key !== key0 && !IX.ALIAS[key0] && !mostrar[key]) mostrar[key] = f.ajustada;
      if (EXCL[key]) return;
      if (resg[key]) resguardos[key] = (resguardos[key] || 0) + 1; else casos[key] = (casos[key] || 0) + 1;
    });
    Object.keys(urbanas).forEach(function (k) {            // nombre mostrado: la grafía más usada (a igualdad, la menor en orden alfabético)
      var sp = nomUrb[k], best = null;
      Object.keys(sp).sort().forEach(function (n) { if (best === null || sp[n] > sp[best]) best = n; });
      urbanas[k].nombre = best;
    });
    return { casos: casos, resguardos: resguardos, mostrar: mostrar, urbanas: urbanas, total_registros: total };
  }

  // ---------------------------------------------------------------- novedades por corte (recuadro «Novedades» del mapa)
  // Un corte se nombra «SE 44 de 2026». Las semanas epidemiológicas van de domingo a sábado; la SE 1 es la que deja 4 o más días en enero.
  function semanasEpi(anio) {
    function ini(a) { var e = Date.UTC(a, 0, 1), d = new Date(e).getUTCDay(), s = e - d * 864e5; if (d > 3) s += 7 * 864e5; return s; }
    return Math.round((ini(anio + 1) - ini(anio)) / (7 * 864e5));
  }
  function leerCorte(txt) { var m = /SE\s*(\d{1,2})\D+(\d{4})/i.exec(String(txt || '')); return m ? { se: parseInt(m[1], 10), anio: parseInt(m[2], 10) } : null; }
  function ordenCorte(txt) { var c = leerCorte(txt); return c ? c.anio * 100 + c.se : 0; }
  // Semanas que cubre un corte desde el anterior: «SE 33 a 44 de 2026»
  function rangoCorte(corte, anterior) {
    var c = leerCorte(corte), a = leerCorte(anterior); if (!c) return String(corte || '');
    if (!a || ordenCorte(anterior) >= ordenCorte(corte)) return 'SE ' + c.se + ' de ' + c.anio;
    var se = a.se + 1, an = a.anio; if (se > semanasEpi(an)) { se = 1; an++; }
    if (an === c.anio) return se === c.se ? 'SE ' + se + ' de ' + an : 'SE ' + se + ' a ' + c.se + ' de ' + an;
    return 'SE ' + se + ' de ' + an + ' a SE ' + c.se + ' de ' + c.anio;
  }
  // cortes: [{corte, anterior, fecha, filas:[{cod, ajustada, entidad}]}] → conteos por lugar de los casos que entraron al mapa en cada corte
  function novedadesDatos(cortes, REF, IX, extra) {
    return cortes.slice().sort(function (x, y) { return ordenCorte(y.corte) - ordenCorte(x.corte); }).map(function (c) {
      var d = generarDatos(c.filas || [], REF, IX, extra), n = 0, o = { corte: c.corte, anterior: c.anterior || '', fecha: c.fecha || '' };
      Object.keys(d.casos).forEach(function (k) { n += d.casos[k]; });
      Object.keys(d.resguardos).forEach(function (k) { n += d.resguardos[k]; });
      Object.keys(d.urbanas).forEach(function (k) { n += d.urbanas[k].n; });
      o.n = n; o.casos = d.casos;
      if (Object.keys(d.resguardos).length) o.resguardos = d.resguardos;
      if (Object.keys(d.urbanas).length) o.urbanas = d.urbanas;
      return o;
    });
  }

  // Contenido de datos.js (solo conteos por vereda; ningún dato de personas)
  function escribirDatosJs(d, meta) {
    function ord(o) { var r = {}; Object.keys(o || {}).sort().forEach(function (k) { r[k] = o[k]; }); return r; }
    function esc(t) { return String(t).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
    var obj = { casos: ord(d.casos), corte: meta.corte };
    if (meta.fecha_corte) obj.fecha_corte = meta.fecha_corte;      // fecha de la última información (AAAA-MM-DD); el visor la muestra en el título
    obj.generado = meta.generado; obj.periodo = meta.periodo; obj.resguardos = ord(d.resguardos); obj.total_registros = d.total_registros;
    if (Object.keys(d.mostrar || {}).length) obj.mostrar = ord(d.mostrar);
    if (Object.keys(d.urbanas || {}).length) obj.urbanas = ord(d.urbanas);
    if (meta.credito && meta.credito.length) obj.credito = meta.credito.map(function (t, i) { return i === 0 ? '<b>' + esc(t) + '</b>' : esc(t); });
    if (d.novedades && d.novedades.length) obj.novedades = d.novedades.map(function (c) {      // casos que entraron al mapa en cada corte (conteos por lugar)
      var o = { corte: c.corte, anterior: c.anterior, fecha: c.fecha, n: c.n, casos: ord(c.casos) };
      if (c.resguardos) o.resguardos = ord(c.resguardos); if (c.urbanas) o.urbanas = ord(c.urbanas);
      return o;
    });
    return '// DATOS DEL MAPA: solo conteos por vereda (sin datos de personas). Lo regenera actualizador.html\nvar DATOS = ' + JSON.stringify(obj) + ';\n';
  }

  return { semanasEpi: semanasEpi, leerCorte: leerCorte, ordenCorte: ordenCorte, rangoCorte: rangoCorte, novedadesDatos: novedadesDatos, parseFecha: parseFecha, fechaMaxima: fechaMaxima, fechaLarga: fechaLarga, ordenFechas: ordenFechas, ENT_URBANA: ENT_URBANA, MIN_CASOS_PIN: MIN_CASOS_PIN, escribirDatosJs: escribirDatosJs, norm: norm, limpiarVereda: limpiarVereda, cyrb53: cyrb53, normId: normId, huellaBase: huellaBase,
           parseCSV: parseCSV, mapearColumnas: mapearColumnas, extraerCaso: extraerCaso, COLS_REQUERIDAS: COLS_REQUERIDAS,
           construirIndices: construirIndices, lev: lev, relacion: relacion, sugerencias: sugerencias,
           otrosMunicipios: otrosMunicipios, clasificar: clasificar, generarDatos: generarDatos };
});
