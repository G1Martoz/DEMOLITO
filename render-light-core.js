/* ================================================================
 * render-light-core.js  ·  DEMO-LITO  ·  Ultra-Light Render Engine
 * Optimizado para: 1-Core @ 1.1 GHz · 4 GB RAM  (Argentina - Ley 20.744)
 * Sin dependencias externas. Vanilla JS puro. Sin loops en hot path.
 * ================================================================ */

'use strict';

/* ==========================================================
 * 1. HELPERS Y CASTEO SEGURO (sin regex pesado)
 * ========================================================== */

function _num(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const s = String(v).trim();
  if (!s) return 0;
  const c = s.replace(/[^\d,.-]/g, '');
  const dot = c.lastIndexOf('.');
  const com = c.lastIndexOf(',');
  let n;
  if (com > dot) {
    n = Number(c.replace(/\./g, '').replace(',', '.'));
  } else {
    n = Number(c.replace(/,/g, ''));
  }
  return isFinite(n) ? n : 0;
}

function _txt(v, fallback) {
  if (v === null || v === undefined) return (fallback || '');
  const s = String(v).trim();
  return s.length ? s : (fallback || '');
}

/* Singleton cacheado Intl.NumberFormat (1 instancia, no 1 por render) */
const _fmtARS = (function() {
  try { return new Intl.NumberFormat('es-AR', { style: 'currency', currency: 'ARS', minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  catch (e) { return { format: function(n) { return '$ ' + (Math.round(_num(n) * 100) / 100).toFixed(2).replace('.', ','); } }; }
})();

function formatearMoneda(n) {
  return _fmtARS.format(_num(n));
}

function formatearMonedaCorta(n, ceroMostrar) {
  const v = _num(n);
  if (!ceroMostrar && Math.abs(v) < 0.005) return '';
  return formatearMoneda(v);
}

/* ==========================================================
 * 2. NÚMEROS A LETRAS  (ITERATIVO. Soporta $0 .. $999.999.999.999,99)
 *    Versión optimizada. Sin diccionarios gigantes.
 * ========================================================== */
(function(){
  const _U = ['','UN','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE','VEINTE','VEINTIÚN','VEINTIDÓS','VEINTITRÉS','VEINTICUATRO','VEINTICINCO','VEINTISÉIS','VEINTISIETE','VEINTIOCHO','VEINTINUEVE'];
  const _D = ['','','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const _C = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS','SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];
  function _centenas(n, and) {
    if (!n) return '';
    if (n === 100) return 'CIEN';
    const c = Math.floor(n / 100);
    const r = n % 100;
    let s = _C[c];
    if (r) {
      if (r <= 29) s += (s ? ' ' : '') + _U[r];
      else {
        const d = Math.floor(r / 10);
        const u = r % 10;
        s += (s ? ' ' : '') + _D[d];
        if (u) s += ' Y ' + _U[u];
      }
    }
    return s;
  }
  function _bloqueMil(n) {
    if (!n) return '';
    if (n === 1) return 'UN MIL';
    const c = Math.floor(n / 1000);
    const r = n % 1000;
    let s = '';
    if (c) {
      if (c === 1) s += 'MIL';
      else s += _centenas(c) + ' MIL';
    }
    if (r) s += (s ? ' ' : '') + _centenas(r);
    return s;
  }
  function _bloqueMillones(n) {
    if (!n) return '';
    const m = Math.floor(n / 1000000);
    const r = n % 1000000;
    let s = '';
    if (m) {
      if (m === 1) s = 'UN MILLÓN';
      else s = _bloqueMil(m) + ' MILLONES';
    }
    const b = _bloqueMil(r);
    if (b) s += (s ? ' ' : '') + b;
    return s;
  }
  window.numeroALetrasPesos = function(n) {
    try {
      const v = Math.max(0, _num(n));
      const enteros = Math.floor(v);
      const decimales = Math.round((v - enteros) * 100);
      let txt = '';
      const miles = enteros % 1000000;
      const millones = Math.floor(enteros / 1000000);
      if (millones) txt = _bloqueMillones(enteros);
      else txt = _bloqueMil(enteros);
      if (!txt) txt = 'CERO';
      const suf = (enteros === 1) ? ' PESO' : ' PESOS';
      const centavos = (decimales < 10 ? '0' : '') + String(decimales);
      return (txt.charAt(0) + txt.slice(1).toLowerCase()) + suf + ' con ' + centavos + '/100';
    } catch (e) {
      try { return formatearMoneda(n) + '  [error conversion letras]'; }
      catch (err) { return 'Error'; }
    }
  };
})();

/* ==========================================================
 * 3. NORMALIZACIÓN DEFENSIVA DE FILA EXCEL CLIENTE
 *    Lookup CASE-INSENSITIVE + normaliza espacios/acentos en keys.
 * ========================================================== */
(function(){
  function _buildKeyMap(row) {
    const m = Object.create(null);
    if (!row) return m;
    const keys = Object.keys(row);
    for (let i = 0; i < keys.length; i++) {
      const k = keys[i];
      const nk = String(k || '').trim().toUpperCase().replace(/\s+/g, ' ').replace(/\u00A0/g, ' ');
      m[nk] = k;
    }
    return m;
  }
  function _gk(row, map, candidatesArr, fallback) {
    for (let i = 0; i < candidatesArr.length; i++) {
      const k = candidatesArr[i].trim().toUpperCase().replace(/\s+/g, ' ');
      const orig = map[k];
      if (orig !== undefined && row[orig] !== undefined && row[orig] !== null && String(row[orig]).trim() !== '') return row[orig];
    }
    return fallback;
  }
  window.normalizarRowCliente = function(row) {
    const m = _buildKeyMap(row);
    // Campos extra identificación / categoría (real Excel cliente)
    const periodoR     = _txt(_gk(row, m, ['PERÍODOR','PERIODO R','PERIODOR','PERÍODO R','PERIODOREF','PERÍODOREF'], ''));
    const nivel        = _txt(_gk(row, m, ['NIVEL'], ''));
    const anexo        = _txt(_gk(row, m, ['ANEXO'], ''));
    const ceic         = _txt(_gk(row, m, ['CEIC'], ''));
    const indices      = _txt(_gk(row, m, ['INDICES'], ''));
    const pctAntig     = _num(_gk(row, m, ['%ANTIG','PORC ANTIG','% ANTIG','ANTIG PORC']));
    const fBaja        = _txt(_gk(row, m, ['F.BAJA','F BAJA','FECHA BAJA'], ''));
    const dDescuento   = _num(_gk(row, m, ['DDESCUENTO','D DESCUENTO','DIAS DESCUENTO']));
    // Haberes remunerativos REALES Excel cliente (35+ cols detectadas)
    const tiempoCompleto = _num(_gk(row, m, ['TIEMPO COMPLETO','TIEMPO COMPL']));
    const estadoDocente  = _num(_gk(row, m, ['ESTADO DOCENTE','EST.DOCENTE','EST DOCENTE']));
    const tituloDocente  = _num(_gk(row, m, ['TÍTULO DOCENTE','TITULO DOCENTE','TIT.DOCENTE','TIT DOCENTE']));
    const residencia     = _num(_gk(row, m, ['RESIDENCIA','RESID']));
    const tareaDiferenciada = _num(_gk(row, m, ['TAREA DIFERENCIADA','TAREA DIF','T.DIFERENCIADA']));
    const adicEdEsp      = _num(_gk(row, m, ['ADIC.ED.ESP','ADIC ED ESP','ADICIONAL EDUCACION ESPECIAL','ADIC EDUC ESP']));
    const recDocente     = _num(_gk(row, m, ['REC. DOCENTE','REC DOCENTE','RECARGO DOCENTE']));
    const otrosR         = _num(_gk(row, m, ['OTROSR','OTROS R','OTROS REMUNERATIVOS','OTROS REMUNERATIVOS 1']));
    const otrosR2        = _num(_gk(row, m, ['OTROSR2','OTROS R2','OTROS REMUNERATIVOS 2','OTROS REM 2']));
    const vacProp        = _num(_gk(row, m, ['VACPROP','VAC PROP','VACACIONES PROPORCIONALES','VACACIONES PROP']));
    const descuentoHab   = _num(_gk(row, m, ['DESCUENTO','DESC'])); // descuento negativo sobre haberes
    // Haberes NO remunerativos REALES
    const salarioFamiliar = _num(_gk(row, m, ['SALARIO FAMILIAR','SAL FAM','ASIGNACIONES FAMILIARES','ASIGNACION FAMILIAR']));
    const asigGobProv     = _num(_gk(row, m, ['ASIGGOBPROV','ASIG GOB PROV','ASIGNACION GOBIERNO PROVINCIAL','GOB PROV']));
    const asigGobNac      = _num(_gk(row, m, ['ASIGGOBNAC','ASIG GOB NAC','ASIGNACION GOBIERNO NACIONAL','GOB NAC']));
    const aula            = _num(_gk(row, m, ['AULA']));
    const otrosNR         = _num(_gk(row, m, ['OTROSNR','OTROS NR','OTROS NO REMUNERATIVOS','OTROS NO REM 1']));
    const otrosNR2        = _num(_gk(row, m, ['OTROSNR2','OTROS NR2','OTROS NO REMUNERATIVOS 2','OTROS NO REM 2']));
    // Contribuciones / totales extra
    const totalAyC        = _num(_gk(row, m, ['TOTAL A Y C','TOTAL A+C','TOTAL APORTES Y CONTRIBUCIONES']));
    const totalAportado   = _num(_gk(row, m, ['TOTALAPORTADO','TOTAL APORTADO','TOTAL PAGADO']));
    // Metadatos referencia
    const refSalFam       = _txt(_gk(row, m, ['REFSALFAM','REF SAL FAM','REFERENCIA SALARIO FAMILIAR'], ''));
    const detalleOtrosR   = _txt(_gk(row, m, ['DETALLE OTROSR','DET OTROS R','DETALLE OTROS REMUNERATIVOS'], ''));
    const detalleOtrosNR  = _txt(_gk(row, m, ['DETALLE OTROSNR','DET OTROS NR','DETALLE OTROS NO REMUNERATIVOS'], ''));
    const detalleOtrosR2  = _txt(_gk(row, m, ['DETALLE OTROSR2','DET OTROS R2'], ''));
    const detalleOtrosNR2 = _txt(_gk(row, m, ['DETALLE OTROSNR2','DET OTROS NR2'], ''));
    const reemplazaA      = _txt(_gk(row, m, ['REEMPLAZA A','REEMPLAZA'], ''));
    const obs             = _txt(_gk(row, m, ['OBS','OBSERVACIONES','OBSERVACIÓN'], ''));

    // --- Totales: FALLBACK por suma de partes si columnas TOTALREM / TOTALNOREM son 0 o faltan ---
    const basico      = _num(_gk(row, m, ['BASICO','BÁSICO','SUELDO BÁSICO','SUELDO BASICO']));
    const presentismo = _num(_gk(row, m, ['PRESENTISMO']));
    const antig       = _num(_gk(row, m, ['ANTIG','ANTIGUEDAD','ANTIGÜEDAD (MONTO)']));
    const sacProp     = _num(_gk(row, m, ['SACPROP','SAC PROP','SAC PROPORCIONAL','AGUINALDO','SAC']));
    let totalRem     = _num(_gk(row, m, ['TOTALREM','TOTAL REM','TOTAL REMUNERATIVO']));
    let totalNoRem   = _num(_gk(row, m, ['TOTALNOREM','TOTAL NO REM','TOTAL NO REMUNERATIVO']));
    const totalNeto  = _num(_gk(row, m, ['TOTALNETO','TOTAL NETO','NETO A COBRAR','NETO']));
    // Aportes
    const apSS       = _num(_gk(row, m, ['AP.SS','AP SS','AP JUB','JUBILACIÓN','JUBILACION','APORTE SEGURIDAD SOCIAL']));
    const apOS       = _num(_gk(row, m, ['AP.OS','AP OS','OBRA SOCIAL TRAB','APORTE OBRA SOCIAL']));
    const apFSP      = _num(_gk(row, m, ['AP.FSP','AP FSP','FAECYS','FSP']));
    const apSindical = _num(_gk(row, m, ['AP.SINDICAL','AP SINDICAL','APORTE SINDICAL','SINDICAL']));
    const apCC       = _num(_gk(row, m, ['AP.CC','AP CC','CAJA COMPLEMENTARIA','SEGURO VIDA']));
    const totalAportes = _num(_gk(row, m, ['TOTAL APORTES','TOTALES APORTES','TOTAL AP.']));
    // Contribuciones patronales
    const contribSS   = _num(_gk(row, m, ['CONTRIBSS','CONTRIB SS','SEGURIDAD SOCIAL EMP','SIPA PATRONAL']));
    const contribOS   = _num(_gk(row, m, ['CONTRIBOS','CONTRIB OS','OBRA SOCIAL EMP']));
    const contribART  = _num(_gk(row, m, ['CONTRIBART','CONTRIB ART','ART + FFEP','ART']));
    const contribSVC  = _num(_gk(row, m, ['CONTRIBSVC','CONTRIB SVC','SEGURO DE VIDA','SVC']));
    const contribFSP  = _num(_gk(row, m, ['CONTRIBFSP','CONTRIB FSP','OSECAC PATRONAL','FSP PATRONAL']));
    const totalContribuciones = _num(_gk(row, m, ['TOTAL CONTRIBUCIONES','TOTAL CONTRIB','TOTALES CONTRIBUCIONES']));

    // === FALLBACK SUMA DE PARTES: si TOTALREM es 0 o muy chico → recalcular ===
    const sumaRem = basico + presentismo + tiempoCompleto + estadoDocente + tituloDocente
                  + residencia + tareaDiferenciada + adicEdEsp + recDocente + antig
                  + otrosR + otrosR2 + vacProp + sacProp - Math.max(0, descuentoHab);
    if (Math.abs(totalRem) < 0.01 && Math.abs(sumaRem) > 0.01) totalRem = sumaRem;
    // === FALLBACK SUMA NO REM ===
    const sumaNoRem = salarioFamiliar + asigGobProv + asigGobNac + aula + otrosNR + otrosNR2;
    if (Math.abs(totalNoRem) < 0.01 && Math.abs(sumaNoRem) > 0.01) totalNoRem = sumaNoRem;

    return {
      // Identificación (incluye COD. SERIALIZADOR con punto para SAC)
      legajo:       _txt(_gk(row, m, ['COD. SERIALIZADOR','CODIGO SERIALIZADOR','CODIGO','COD. SERIAL','COD SERIALIZADOR','LEG','LEGAJO','LEGAJO NRO','N LEGAJO'], '')),
      cuil:         _txt(_gk(row, m, ['CUIL','CUIT'], '')),
      apellido:     _txt(_gk(row, m, ['APELLIDO','APELLIDOS'], '')),
      nombre:       _txt(_gk(row, m, ['NOMBRE','NOMBRES'], '')),
      banco:        _txt(_gk(row, m, ['BANCO'], '')), // puede venir vacío
      // Fechas y categoría EXTRA
      periodo:      _txt(_gk(row, m, ['PERÍODO','PERIODO','MES ANIO','MES/AÑO','MES'], '')),
      periodoR:     periodoR,
      nivel:        nivel,
      anexo:        anexo,
      ceic:         ceic,
      indices:      indices,
      pctAntig:     pctAntig,
      descripcion:  _txt(_gk(row, m, ['DESCRIPCION','DESCRIPCIÓN','CATEGORIA','CATEGORÍA','CARGO'], '')),
      sitRev:       _txt(_gk(row, m, ['SITREV','SIT. REV.','SITUACIÓN DE REVISTA','SITUACION REVISTA'], '')),
      antigAM:      _txt(_gk(row, m, ['ANTIG.A/M','ANTIGÜEDAD','ANTIGUEDAD','ANTIG. A/M','ANTIG'], '')),
      fAlta:        _txt(_gk(row, m, ['F.ALTA','F ALTA','FECHA ALTA','INGRESO'], '')),
      fBaja:        fBaja,
      dias:         _num(_gk(row, m, ['DIAS','DÍAS'])),
      horas:        _num(_gk(row, m, ['HORAS','HS'])),
      dDescuento:   dDescuento,
      // === HABERES REMUNERATIVOS REALES (14 filas HTML) ===
      basico:            basico,
      presentismo:       presentismo,
      tiempoCompleto:    tiempoCompleto,
      estadoDocente:     estadoDocente,
      tituloDocente:     tituloDocente,
      residencia:        residencia,
      tareaDiferenciada: tareaDiferenciada,
      adicEdEsp:         adicEdEsp,
      recDocente:        recDocente,
      antig:             antig,
      otrosR:            otrosR,
      otrosR2:           otrosR2,
      vacProp:           vacProp,
      sacProp:           sacProp,
      descuentoHab:      descuentoHab, // puede ser 0
      // === HABERES NO REMUNERATIVOS REALES (6 filas HTML) ===
      salarioFamiliar:   salarioFamiliar,
      asigGobProv:       asigGobProv,
      asigGobNac:        asigGobNac,
      aula:              aula,
      otrosNR:           otrosNR,
      otrosNR2:          otrosNR2,
      // Totales (con fallback por suma)
      totalRem:     totalRem,
      totalNoRem:   totalNoRem,
      totalNeto:    totalNeto,
      // Deducciones / Aportes trabajador (5 filas + subtotal)
      apSS:         apSS,
      apOS:         apOS,
      apFSP:        apFSP,
      apSindical:   apSindical,
      apCC:         apCC,
      totalAportes: Math.abs(totalAportes) > 0.01 ? totalAportes : (apSS+apOS+apFSP+apSindical+apCC),
      // Contribuciones Patronales (6 filas)
      contribSS:    contribSS,
      contribOS:    contribOS,
      contribART:   contribART,
      contribSVC:   contribSVC,
      contribFSP:   contribFSP,
      totalContribuciones: Math.abs(totalContribuciones) > 0.01 ? totalContribuciones : (contribSS+contribOS+contribART+contribSVC+contribFSP),
      totalAyC:     totalAyC,
      totalAportado:totalAportado,
      // Metadatos (referencia, detalles, obs)
      refSalFam:    refSalFam,
      detalleOtrosR: detalleOtrosR,
      detalleOtrosNR: detalleOtrosNR,
      detalleOtrosR2: detalleOtrosR2,
      detalleOtrosNR2: detalleOtrosNR2,
      reemplazaA:   reemplazaA,
      obs:          obs
    };
  };
})();

/* ==========================================================
 * 4. RENDER DEL RECIBO — ULTRA DIRECTO.
 *    ~70 textContent. SIN loops, sin innerHTML, sin forEach, sin Object.keys.
 *    Si la celda/id no existe → no falla (void-0 chaining por getElementById).
 * ========================================================== */
window.renderReciboDirecto = function(index) {
  const emp = window.empleadosData ? window.empleadosData[index] : null;
  if (!emp) return;

  // --------- Cabecera recibo ---------
  const rNombre    = document.getElementById('rec-nombre');
  const rLegajo    = document.getElementById('rec-legajo');
  const rBruto     = document.getElementById('rec-bruto');
  const rBruto2    = document.getElementById('rec-bruto-2');
  const rAntig     = document.getElementById('rec-antiguedad');
  const rIngreso   = document.getElementById('rec-ingreso');
  const rCat       = document.getElementById('rec-categoria');
  const rCuil      = document.getElementById('rec-cuil');
  const rBanco     = document.getElementById('rec-banco');
  const rPerAport  = document.getElementById('rec-periodo-aportes');
  const rFchAport  = document.getElementById('rec-fecha-aportes');
  const rMesAno    = document.getElementById('rec-mes-ano');
  const rTituloLiq = document.getElementById('rec-liquidacion-titulo');

  const nombreCompleto = emp.apellido || emp.nombre ? (emp.apellido + (emp.apellido && emp.nombre ? ', ' : '') + emp.nombre) : '—';
  if (rNombre)    rNombre.textContent = nombreCompleto.toUpperCase();
  if (rLegajo)    rLegajo.textContent  = emp.legajo  || '—';
  const brutoCalc = (emp.totalRem || 0) + (emp.totalNoRem || 0);
  const _sumaAportesInd = (emp.apSS||0)+(emp.apOS||0)+(emp.apFSP||0)+(emp.apSindical||0)+(emp.apCC||0);
  const bruto = (emp.totalNeto||0) + (emp.totalAportes || _sumaAportesInd);
  const brutoMostrar = emp.basico ? (brutoCalc || bruto || 0) : (bruto || brutoCalc || 0);
  const brutoMostrarFin = brutoMostrar || 0;
  if (rBruto)     rBruto.textContent   = formatearMoneda(brutoMostrarFin);
  if (rBruto2)    rBruto2.textContent  = formatearMoneda(brutoMostrarFin);
  if (rAntig)     rAntig.textContent    = emp.antigAM || (emp.antig ? (emp.antig.toFixed ? (Math.round(emp.antig*100)/100).toFixed(2) : String(emp.antig)) : '—');
  let ingresoTxt = emp.fAlta || '—';
  if (emp.fBaja) ingresoTxt = (emp.fAlta||'—') + ' → Baja: ' + emp.fBaja;
  if (rIngreso)   rIngreso.textContent  = ingresoTxt;
  if (rCat)       rCat.textContent      = emp.descripcion || emp.sitRev || '—';
  if (rCuil)      rCuil.textContent     = emp.cuil || '—';
  if (rBanco)     rBanco.textContent    = emp.banco || 'Según nómina';
  if (rPerAport)  rPerAport.textContent = emp.periodo || emp.periodoR || '—';
  if (rFchAport)  rFchAport.textContent = (typeof window === 'object' && window.inputPeriodoFecha) ? window.inputPeriodoFecha : emp.periodo || emp.periodoR || '—';
  if (rMesAno)    rMesAno.textContent   = (typeof window === 'object' && window.getPeriodLabel) ? window.getPeriodLabel() : (emp.periodo || emp.periodoR || '—');
  if (rTituloLiq) {
    const partes = [];
    if (emp.periodoR || emp.periodo) partes.push('Período ' + (emp.periodoR || emp.periodo));
    if (emp.sacProp > 0.01 || /SAC/i.test(emp.periodo||'') || /SAC/i.test(emp.periodoR||'')) partes[0] += ' · SAC';
    if (emp.descripcion) partes.push(emp.descripcion);
    if (emp.sitRev) partes.push(emp.sitRev);
    if (emp.nivel) partes.push('Nivel ' + emp.nivel);
    if (emp.anexo) partes.push('Anexo ' + emp.anexo);
    if (emp.horas) partes.push(emp.horas + 'h');
    rTituloLiq.textContent = 'LIQUIDACIÓN HABERES' + (partes.length ? ' — ' + partes.join(' · ') : '');
  }

  // --------- Tabla contribuciones patronales (FILAS FIJAS. Sin innerHTML) ---------
  const cc_unidad = 'MONTO';
  function _filaContrib(prefijo, concepto, unidad, base, monto) {
    const c = document.getElementById(prefijo+'-c'); if (c) c.textContent = concepto;
    const u = document.getElementById(prefijo+'-u'); if (u) u.textContent = unidad || cc_unidad;
    const b = document.getElementById(prefijo+'-b'); if (b) b.textContent = formatearMonedaCorta(base, false);
    const m = document.getElementById(prefijo+'-m'); if (m) m.textContent = formatearMonedaCorta(monto, false);
  }
  const baseContrib = brutoMostrar;
  _filaContrib('contrib-1', 'Aportes Seguridad Social (Jubilación / SIPA)',        '% / MONTO', baseContrib, emp.contribSS);
  _filaContrib('contrib-2', 'Contribución Obra Social (Empleador)',                 '% / MONTO', baseContrib, emp.contribOS);
  _filaContrib('contrib-3', 'ART + FFEP / Aseguradora Riesgos Trabajo',            'MONTO',     baseContrib, emp.contribART);
  _filaContrib('contrib-4', 'Seguro de Vida Obligatorio + CCT (SVC)',              'MONTO',     baseContrib, emp.contribSVC);
  _filaContrib('contrib-5', 'OSECAC / Fondo Sindical Patronal (FSP)',              '% / MONTO', baseContrib, emp.contribFSP);
  _filaContrib('contrib-6', 'Aportes INACAP y otros (según nómina)',                'MONTO',     0,           0);
  const rSubContrib = document.getElementById('rec-subtotal-contrib');
  const totalContrib = emp.totalContribuciones || (emp.contribSS+emp.contribOS+emp.contribART+emp.contribSVC+emp.contribFSP);
  if (rSubContrib) rSubContrib.textContent = formatearMoneda(totalContrib);
  const rCostoTotal = document.getElementById('rec-costo-total');
  if (rCostoTotal) rCostoTotal.textContent = formatearMoneda(brutoMostrar + totalContrib);

  // --------- Tabla conceptos (HABERES FIJOS + APORTES FIJOS. Sin innerHTML) ---------
  function _filaConc(prefijo, concepto, unidad, base, rem, norem, desc) {
    const c = document.getElementById(prefijo+'-c'); if (c) c.textContent = concepto;
    const u = document.getElementById(prefijo+'-u'); if (u) u.textContent = unidad || '—';
    const b = document.getElementById(prefijo+'-b'); if (b) b.textContent = formatearMonedaCorta(base, false);
    const r = document.getElementById(prefijo+'-r'); if (r) r.textContent = formatearMonedaCorta(rem, false);
    const nr= document.getElementById(prefijo+'-nr');if (nr)nr.textContent = formatearMonedaCorta(norem, false);
    const d = document.getElementById(prefijo+'-d'); if (d) d.textContent = formatearMonedaCorta(desc, false);
  }
  const dias = emp.dias || 30;
  const horas = emp.horas || '';
  const unidadTiempo = horas ? (String(horas)+' h') : (String(dias)+' d');
  // === RECÁLCULO POR SUMA DE PARTES (por si TotalRem Excel no coincidía con columnas individuales) ===
  const _sumaRemCalc = emp.basico + emp.presentismo + emp.tiempoCompleto + emp.estadoDocente
                    + emp.tituloDocente + emp.residencia + emp.tareaDiferenciada + emp.adicEdEsp
                    + emp.recDocente + emp.antig + emp.otrosR + emp.otrosR2 + emp.vacProp + emp.sacProp
                    - Math.max(0, emp.descuentoHab || 0);
  const TotalRemMostrar = Math.abs(emp.totalRem) > 0.01 ? emp.totalRem : _sumaRemCalc;
  const _sumaNoRemCalc = emp.salarioFamiliar + emp.asigGobProv + emp.asigGobNac + emp.aula + emp.otrosNR + emp.otrosNR2;
  const TotalNoRemMostrar = Math.abs(emp.totalNoRem) > 0.01 ? emp.totalNoRem : _sumaNoRemCalc;
  const totalAportesCalc = emp.totalAportes || (emp.apSS+emp.apOS+emp.apFSP+emp.apSindical+emp.apCC);
  // === HABERES REMUNERATIVOS (14 filas: conc-1..conc-14) ===
  _filaConc('conc-1',  'Sueldo Básico',                               unidadTiempo, emp.basico,            emp.basico,            0, 0);
  _filaConc('conc-2',  'Presentismo',                                 '% / MONTO',  emp.basico,            emp.presentismo,       0, 0);
  _filaConc('conc-3',  'Tiempo Completo',                             'MONTO',      0,                     emp.tiempoCompleto,    0, 0);
  _filaConc('conc-4',  'Estado Docente',                              'MONTO',      0,                     emp.estadoDocente,     0, 0);
  _filaConc('conc-5',  'Título Docente',                              'MONTO',      0,                     emp.tituloDocente,     0, 0);
  _filaConc('conc-6',  'Residencia',                                  'MONTO',      0,                     emp.residencia,        0, 0);
  _filaConc('conc-7',  'Tarea Diferenciada',                          'MONTO',      0,                     emp.tareaDiferenciada, 0, 0);
  _filaConc('conc-8',  'Adicional Educación Especial (Adic.Ed.Esp)',  'MONTO',      0,                     emp.adicEdEsp,         0, 0);
  _filaConc('conc-9',  'Recargo Docente (Rec. Docente)',              'MONTO',      0,                     emp.recDocente,        0, 0);
  _filaConc('conc-10', 'Antigüedad (Haberes)',                        '% / MONTO',  emp.basico,            emp.antig,             0, 0);
  _filaConc('conc-11', 'Otros Remunerativos (OtrosR)',                'MONTO',      0,                     emp.otrosR,            0, 0);
  _filaConc('conc-12', 'Otros Remunerativos 2 (OtrosR2)',             'MONTO',      0,                     emp.otrosR2,           0, 0);
  _filaConc('conc-13', 'Vacaciones Proporcionales (VacProp)',         'Días/MONTO', 0,                     emp.vacProp,           0, 0);
  _filaConc('conc-14', 'SAC Proporcional / Aguinaldo (SacProp)',      'Días/MONTO', 0,                     emp.sacProp,           0, 0);
  // === SUBTOTAL REMUNERATIVO (conc-15) ===
  _filaConc('conc-15', 'SUB TOTAL REMUNERATIVO',                      'Subtotal',   TotalRemMostrar,       TotalRemMostrar,       0, 0);
  // === HABERES NO REMUNERATIVOS (6 filas: conc-16..conc-21) ===
  _filaConc('conc-16', 'Salario Familiar (No Remunerativo)',          'MONTO',      0,                     0,                     emp.salarioFamiliar, 0);
  _filaConc('conc-17', 'Asignación Gobierno Provincial (AsigGobProv)','MONTO',      0,                     0,                     emp.asigGobProv,     0);
  _filaConc('conc-18', 'Asignación Gobierno Nacional (AsigGobNac)',   'MONTO',      0,                     0,                     emp.asigGobNac,      0);
  _filaConc('conc-19', 'Aula (No Remunerativo)',                      'MONTO',      0,                     0,                     emp.aula,            0);
  _filaConc('conc-20', 'Otros No Remunerativos (OtrosNR)',            'MONTO',      0,                     0,                     emp.otrosNR,         0);
  _filaConc('conc-21', 'Otros No Remunerativos 2 (OtrosNR2)',         'MONTO',      0,                     0,                     emp.otrosNR2,        0);
  // === SUBTOTAL NO REMUNERATIVO (conc-22) ===
  _filaConc('conc-22', 'SUB TOTAL NO REMUNERATIVO',                   'Subtotal',   TotalNoRemMostrar,     0,                     TotalNoRemMostrar,   0);
  // === DESCUENTOS / APORTES TRABAJADOR (5 filas: conc-23..conc-27) ===
  _filaConc('conc-23', 'Aporte Jubilatorio 11% / SS (Ap.SS)',         '11%',        baseContrib,           0,                     0,                   emp.apSS);
  _filaConc('conc-24', 'Aporte Obra Social 3% (Ap.OS)',               '3%',         baseContrib,           0,                     0,                   emp.apOS);
  _filaConc('conc-25', 'FSP / FAECYS (Ap.FSP)',                       '%',          baseContrib,           0,                     0,                   emp.apFSP);
  _filaConc('conc-26', 'Aporte Sindical (Ap.Sindical)',               '%',          baseContrib,           0,                     0,                   emp.apSindical);
  _filaConc('conc-27', 'Caja Complementaria / Seguro Vida (Ap.CC)',   'MONTO',      baseContrib,           0,                     0,                   emp.apCC);
  // === SUBTOTAL DESCUENTOS (conc-28) ===
  _filaConc('conc-28', 'SUB TOTAL DESCUENTOS',                        'Subtotal',   totalAportesCalc,      0,                     0,                   totalAportesCalc);
  // -------- Recalcular brutoMostrar y netoCalc CON FALLBACK SUMA (mostrado re-render cabecera bruto --------
  const brutoCalcFinal = TotalRemMostrar + TotalNoRemMostrar;
  const brutoMostrarFinal = emp.basico ? (brutoCalcFinal || bruto || 0) : (bruto || brutoCalcFinal || 0);
  if (rBruto)  rBruto.textContent  = formatearMoneda(brutoMostrarFinal);
  if (rBruto2) rBruto2.textContent = formatearMoneda(brutoMostrarFinal);
  // -------- Totales finales línea fina (usar variables CON FALLBACK SUMA --------
  const tRem = document.getElementById('rec-total-rem');
  const tNoRem = document.getElementById('rec-total-norem');
  const tDesc = document.getElementById('rec-total-desc');
  const tNeto = document.getElementById('rec-sueldo-neto');
  const netoCalc = (Math.abs(emp.totalNeto) > 0.01) ? emp.totalNeto : (TotalRemMostrar + TotalNoRemMostrar - totalAportesCalc);
  if (tRem)    tRem.textContent    = formatearMoneda(TotalRemMostrar);
  if (tNoRem)  tNoRem.textContent  = formatearMoneda(TotalNoRemMostrar);
  if (tDesc)   tDesc.textContent   = formatearMoneda(totalAportesCalc);
  if (tNeto)   tNeto.textContent   = formatearMoneda(Math.max(0, netoCalc));
  const nletras = document.getElementById('rec-neto-letras');
  if (nletras) nletras.textContent = window.numeroALetrasPesos(Math.max(0, netoCalc));
  const rBD = document.getElementById('rec-banco-deposito');
  const rCD = document.getElementById('rec-cuenta-deposito');
  if (rBD) rBD.textContent = emp.banco || 'Según nómina';
  if (rCD) rCD.textContent = emp.cuentaDeposito || emp.cuenta || 'Según nómina';

  // --------- Actualización PERIODO DEL RECIBO texto repetido Lugar y fecha (si existe hook) ---------
  if (typeof window === 'object' && typeof window.updatePeriodInRecibo === 'function') {
    try { window.updatePeriodInRecibo(document.getElementById('input-periodo') ? document.getElementById('input-periodo').value : ''); } catch (e) {}
  }

  // --------- Distribución costo TOTAL + TORTA SVG DONUT (0 librerías) ---------
  const costoObj = {
    neto:       Math.max(0, netoCalc),
    jub:        Math.max(0, emp.contribSS + (emp.apSS || 0)),
    obra:       Math.max(0, emp.contribOS + (emp.apOS || 0)),
    pami:       0,
    sind:       Math.max(0, (emp.apSindical || 0) + emp.contribFSP),
    art:        Math.max(0, emp.contribART)
  };
  const totalSum = costoObj.neto + costoObj.jub + costoObj.obra + costoObj.pami + costoObj.sind + costoObj.art;
  const denom = Math.max(0.01, totalSum);
  const _pct = v => Math.max(0, Math.min(100, Math.round((v * 1000) / denom) / 10));
  const distRows = [
    { pref: 'dist-1', label: 'Sueldo Neto',                valor: costoObj.neto, color: '#2563eb' },
    { pref: 'dist-2', label: 'Jubilación (SIPA + aporte)', valor: costoObj.jub,  color: '#dc2626' },
    { pref: 'dist-3', label: 'Obra Social (Emp + Trab)',   valor: costoObj.obra, color: '#16a34a' },
    { pref: 'dist-4', label: 'INSSJP / PAMI',              valor: costoObj.pami, color: '#ea580c' },
    { pref: 'dist-5', label: 'Costo Sindical (Sind + FSP)', valor: costoObj.sind, color: '#f59e0b' },
    { pref: 'dist-6', label: 'ART + Seguros',               valor: costoObj.art,  color: '#0891b2' }
  ];
  const DONUT_CIRC = 439.822971502571; // 2 * Math.PI * 70 (r=70 viewBox 200x200 stroke=30)
  let acumDash = 0;
  for (let i = 0; i < distRows.length; i++) {
    const x = distRows[i];
    const elLbl = document.getElementById(x.pref + '-lbl');
    const elVal = document.getElementById(x.pref + '-val');
    const elPct = document.getElementById(x.pref + '-pct');
    const elCirc = document.getElementById('dist-svg-circle-' + (i + 1));
    const pct = _pct(x.valor);
    const dash = (x.valor > 0 && isFinite(x.valor)) ? Math.max(0, Math.min(DONUT_CIRC, (x.valor / denom) * DONUT_CIRC)) : 0;
    if (elLbl) elLbl.textContent = x.label;
    if (elVal) elVal.textContent = formatearMoneda(x.valor);
    if (elPct) elPct.textContent = pct.toFixed(1) + ' %';
    if (elCirc) {
      elCirc.setAttribute('stroke-dasharray', (dash.toFixed(3)) + ' ' + (DONUT_CIRC.toFixed(3)));
      elCirc.setAttribute('stroke-dashoffset', (-acumDash).toFixed(3));
    }
    acumDash += dash;
  }
  // Texto central donut
  const dcNeto = document.getElementById('dist-center-neto');
  if (dcNeto) dcNeto.textContent = formatearMoneda(netoCalc);
  const dcPct = document.getElementById('dist-center-pct');
  if (dcPct) dcPct.textContent = _pct(costoObj.neto).toFixed(1) + ' %';
  // Actualizar detalles span (det-sindical-total etc. si existen)
  const pairs = [
    ['det-sindical-total', costoObj.sind],
    ['det-sindical-emp',   emp.contribFSP],
    ['det-sindical-trab',  emp.apSindical],
    ['det-seg-social-total', costoObj.jub],
    ['det-seg-social-emp',   emp.contribSS],
    ['det-seg-social-trab',  emp.apSS],
    ['det-inssjp-total',   costoObj.pami],
    ['det-inssjp-emp',     0],
    ['det-inssjp-trab',    0],
    ['det-obra-social-total', costoObj.obra],
    ['det-obra-social-emp',   emp.contribOS],
    ['det-obra-social-trab',  emp.apOS],
    ['det-art-total',        costoObj.art],
    ['det-seg-social-inssjp-emp', 0],
    ['det-seg-social-inssjp-trab',0]
  ];
  for (let j = 0; j < pairs.length; j++) {
    const p = pairs[j];
    const el = document.getElementById(p[0]);
    if (el) el.textContent = formatearMoneda(p[1]);
  }
};
