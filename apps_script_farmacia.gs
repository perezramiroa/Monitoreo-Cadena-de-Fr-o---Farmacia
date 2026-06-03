/**
 * ============================================================
 *  VICUS - FARMACIA  |  Google Apps Script (PROFESIONAL V4)
 *  Hospital Natalio Burd
 *  Desarrollado por Ingeniero Perez Ramiro
 * ============================================================
 */

const CONFIG_FARMACIA = {
  folderPDF: '1Kd-8NUFiWCVuiu4enxDUDdgfX0UmFLvu',
  sheetId: '1Lr3mMWnIaU9PZsDQ_kIUDp26bJcLG0BrOJtMPYnmuio',
  alertasFolder: '1TqEPuXuExj6KurrTMK2jpj1Vuvr8fqht'
};

const SENSORES = [
  { id: '2982082', k: 'QNVRZIG08IBMZD2A', n: 'Depósito Farmacia 1', eq: 'Presvac 1 NHC12554', field: 'field1' },
  { id: '2982085', k: 'F7X0FHMTQGVZNHJ4', n: 'Depósito Farmacia 2', eq: 'Presvac 2 NHC13479', field: 'field1' },
  { id: '2982085', k: 'F7X0FHMTQGVZNHJ4', n: 'Depósito Farmacia 2', eq: 'Pico NHC12771', field: 'field2' },
  { id: '2981109', k: 'QQFIB5N64K8KH86H', n: 'Farmacia Ambulatoria', eq: 'Gafa Mini Visu NHC0807', field: 'field1' }
];

function ejecutarReporteSemanal() {
  const hoy = new Date();
  const haceSieteDias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const fechaEmision = Utilities.formatDate(hoy, "GMT-3", "dd/MM/yyyy");
  const rangoTexto = Utilities.formatDate(haceSieteDias, "GMT-3", "dd/MM/yyyy") + " - " + fechaEmision;

  // Recolectar todos los feeds para el Sheet semanal
  const feedsPorSensor = [];

  SENSORES.forEach(s => {
    try {
      const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
      if (!data || !data.feeds || data.feeds.length === 0) return;

      // Guardar feeds para el sheet
      feedsPorSensor.push({ sensor: s, feeds: data.feeds });

      const trazabilidad = "AUTO-SEM-" + Utilities.formatDate(hoy, "GMT-3", "yyyyMMdd") + "-" + s.id;
      const analizada = analizarDatos(data.feeds, s.field);
      const conectividad = analizarConectividad(data.feeds, s.field);
      const grafico = generarGraficoCurva(data.feeds, s.field, s.n);

      const pdfBlob = generarPDFOficial(s, fechaEmision, rangoTexto, trazabilidad, analizada, conectividad, grafico);
      
      const carpeta = DriveApp.getFolderById(CONFIG_FARMACIA.folderPDF);
      const file = carpeta.createFile(pdfBlob);
      file.setName("Informe_Oficial_" + s.n.replace(/ /g,"_") + "_" + trazabilidad + ".pdf");
    } catch (e) {
      console.error("Error en " + s.n + ": " + e.message);
    }
  });

  // Generar Sheet semanal con todos los sensores
  try {
    generarSheetSemanal(feedsPorSensor, rangoTexto, hoy);
  } catch (e) {
    console.error("Error al generar Sheet semanal: " + e.message);
  }
}

function generarPDFOficial(sensor, fecha, rango, trazabilidad, analizada, conectividad, grafico) {
  const doc = DocumentApp.create('Temp_Reporte_' + sensor.n);
  const body = doc.getBody();

  // Márgenes mínimos para maximizar espacio
  body.setMarginLeft(20).setMarginRight(20).setMarginTop(20).setMarginBottom(20);
  const anchoMax = 555; // 595 - 40

  // --- CABECERA (aparece en TODAS las páginas) ---
  const logo = buscarLogoEnDrive("logo_rih.jpg");
  const header = doc.addHeader();
  if (logo) {
    const hp = header.appendParagraph("");
    hp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    hp.appendInlineImage(logo).setWidth(anchoMax).setHeight(60);
  }
  header.appendHorizontalRule();

  // --- TÍTULO ---
  const t1 = body.appendParagraph("INFORME TÉCNICO DE CADENA DE FRÍO - MEDICAMENTOS REFRIGERADOS");
  t1.setFontSize(14).setBold(true).setForegroundColor("#00384d").setSpacingAfter(4);

  // --- NORMATIVA ---
  body.appendParagraph("Según Disposición ANMAT 10.872/2020 y Disposición ANMAT 2069/2018")
    .setFontSize(9).setItalic(true).setSpacingAfter(10);

  // --- DATOS DEL DISPOSITIVO ---
  body.appendParagraph("Dispositivo: " + sensor.n)
    .setBold(true).setFontSize(11).setSpacingBefore(0).setSpacingAfter(2);
  body.appendParagraph("Equipo/Artefacto: " + sensor.eq)
    .setItalic(true).setFontSize(10).setSpacingAfter(2);
  body.appendParagraph("Período: " + rango)
    .setBold(true).setFontSize(10).setSpacingAfter(2);
  body.appendParagraph("Emisión: " + fecha + " | Trazabilidad: " + trazabilidad)
    .setFontSize(8).setSpacingAfter(10);

  // --- GRÁFICO ---
  body.appendParagraph("CURVA TÉRMICA SEMANAL")
    .setBold(true).setFontSize(10).setSpacingBefore(4).setSpacingAfter(4);
  const pChart = body.appendParagraph("");
  pChart.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pChart.appendInlineImage(grafico).setWidth(anchoMax).setHeight(300);

  // --- TABLA DE ALERTAS ---
  body.appendParagraph("ALERTAS Y RECUPERACIONES (2°C - 8°C)")
    .setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaAlertas = [["Fecha y Hora", "Valor", "Estado", "Duración", "Pico Registrado"]];
  if (analizada.alertasFilas.length > 0) {
    analizada.alertasFilas.forEach(f => tablaAlertas.push([f.h, f.v, f.e, f.d, f.p || "--"]));
  } else {
    tablaAlertas.push(["-", "-", "Sin eventos fuera de rango", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaAlertas));

  // --- EVENTOS DE CONECTIVIDAD ---
  body.appendParagraph("EVENTOS DETECTADOS (>10 min sin datos)")
    .setBold(true).setFontSize(10).setSpacingAfter(4);
  const tablaWifi = [["Inicio", "Fin", "Tipo de Corte", "T. Antes", "T. Desp.", "Duración"]];
  if (conectividad.filas.length > 0) {
    conectividad.filas.forEach(f => tablaWifi.push([f.inicio, f.fin, f.tipo, f.antes, f.despues, f.duracion]));
  } else {
    tablaWifi.push(["-", "-", "✅ Sin interrupciones significativas", "-", "-", "-"]);
  }
  estilizarTabla(body.appendTable(tablaWifi));

  // --- ANÁLISIS Y RECOMENDACIONES ---
  body.appendParagraph("ANÁLISIS TÉCNICO:").setBold(true).setFontSize(10);
  if (analizada.textoAnalisis) {
    body.appendParagraph(analizada.textoAnalisis).setFontSize(9).setItalic(true);
  }
  if (conectividad.analisis && conectividad.analisis !== "Sin problemas de conectividad.") {
    body.appendParagraph("Conectividad:").setBold(true).setFontSize(9);
    body.appendParagraph(conectividad.analisis).setFontSize(9).setItalic(true);
  }
  body.appendParagraph("RECOMENDACIONES:").setBold(true).setFontSize(10).setForegroundColor("#00384d");
  body.appendParagraph(analizada.textoRecom + "\n• " + conectividad.recom).setFontSize(9);

  // Nota técnica final
  body.appendParagraph("NOTA TÉCNICA:").setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaTecnica).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  // Nota de responsabilidad
  body.appendParagraph("RESPONSABILIDAD:").setBold(true).setFontSize(9).setForegroundColor("#475569");
  body.appendParagraph(analizada.notaResponsabilidad).setFontSize(8).setItalic(true).setForegroundColor("#475569");

  // --- PIE DE PÁGINA con línea separadora arriba ---
  const footer = doc.addFooter();
  // Línea separadora al inicio del footer
  footer.appendHorizontalRule();
  const logoF = buscarLogoEnDrive("footer.jpg");
  if (logoF) {
    const fp = footer.appendParagraph("");
    fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    fp.appendInlineImage(logoF).setWidth(anchoMax).setHeight(50);
  }

  doc.saveAndClose();
  const pdf = doc.getAs('application/pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

function analizarConectividad(feeds, field) {
  const filas = [];
  let totalMinutos = 0;
  for (let i = 1; i < feeds.length; i++) {
    const d1 = new Date(feeds[i-1].created_at);
    const d2 = new Date(feeds[i].created_at);
    const diff = (d2 - d1) / 60000;
    if (diff > 10) {
      const v1 = parseFloat(feeds[i-1][field]);
      const v2 = parseFloat(feeds[i][field]);
      const tipo = (!isNaN(v2) && v2 !== -127 && (v2 > 8 || v2 < 2)) ? 'Corte Energía' : 'Corte WiFi';
      filas.push({
        inicio:  fmtFecha(d1),
        fin:     fmtFecha(d2),
        tipo:    tipo,
        antes:   (isNaN(v1) || v1 === -127) ? "--" : v1.toFixed(1) + "°C",
        despues: (isNaN(v2) || v2 === -127) ? "--" : v2.toFixed(1) + "°C",
        duracion: formatDur(diff)
      });
      totalMinutos += diff;
    }
  }
  return {
    filas: filas,
    analisis: filas.length > 0
      ? `Se detectaron ${filas.length} ${filas.length === 1 ? 'interrupción' : 'interrupciones'} de datos.\n• Tiempo total sin datos: ${formatDur(totalMinutos)}\n• Durante los cortes no se puede garantizar el control de la cadena de frío.`
      : "Transmisión continua de datos WiFi confirmada. No se registraron brechas en el período.",
    recom: filas.length > 0
      ? "• Verificar el estado del router y la conexión a internet.\n• Revisar la distancia entre el sensor y el punto de acceso WiFi.\n• Considerar registro manual de temperatura durante los períodos sin datos.\n• Evaluar instalación de UPS para el equipo de red."
      : "Mantener el equipo de red en condiciones óptimas para asegurar monitoreo continuo."
  };
}

function analizarDatos(feeds, field) {
  let alertasFilas = [];
  let lastState = 'normal';
  let startTime = null;
  let stats = [];
  let picoValor = null;
  let picoHora  = null;

  feeds.forEach(f => {
    const val = parseFloat(f[field]);
    if (isNaN(val) || val === -127) return;
    const state = (val > 8.0) ? 'Alta' : (val < 2.0) ? 'Baja' : 'normal';

    // Actualizar pico dentro de la alerta activa
    if (state !== 'normal') {
      if (picoValor === null) {
        picoValor = val; picoHora = f.created_at;
      } else {
        if (state === 'Alta' && val > picoValor) { picoValor = val; picoHora = f.created_at; }
        if (state === 'Baja' && val < picoValor) { picoValor = val; picoHora = f.created_at; }
      }
    }

    if (state !== lastState) {
      const hora = fmtFecha(new Date(f.created_at));
      if (state !== 'normal') {
        startTime = new Date(f.created_at);
        picoValor = val; picoHora = f.created_at;
        const label = state === 'Alta' ? "⚠ Alerta Alta (>8°C)" : "❄ Alerta Baja (<2°C)";
        alertasFilas.push({ h: hora, v: val.toFixed(1) + "°C", e: label, d: "--", p: "--" });
      } else if (startTime) {
        const dur = (new Date(f.created_at) - startTime) / 60000;
        const picoStr = picoValor !== null ? picoValor.toFixed(1) + "°C (" + fmtFecha(new Date(picoHora)) + ")" : "--";
        // Actualizar el pico en la fila de inicio de alerta
        if (alertasFilas.length > 0) alertasFilas[alertasFilas.length - 1].p = picoStr;
        alertasFilas.push({ h: hora, v: val.toFixed(1) + "°C", e: "✅ Recuperación", d: formatDur(dur), p: picoStr });
        stats.push({ s: lastState, d: dur });
        picoValor = null; picoHora = null;
      }
      lastState = state;
    }
  });

  const tieneAltas = stats.some(s => s.s === 'Alta');
  const tieneBajas = stats.some(s => s.s === 'Baja');
  const durTotal = stats.reduce((a, b) => a + b.d, 0);

  let textoAnalisis = "Estabilidad térmica confirmada. Sin desvíos en el período.";
  let textoRecom = "• Continuar monitoreo habitual.\n• Realizar mantenimiento preventivo según calendario.\n• Verificar calibración del sensor periódicamente.";

  if (stats.length > 0) {
    textoAnalisis = `Se detectaron ${stats.length} ${stats.length === 1 ? 'desvío térmico' : 'desvíos térmicos'} (duración total: ${formatDur(durTotal)}).\n`;
    if (tieneAltas) textoAnalisis += "• Temperatura ALTA (>8°C): riesgo de degradación de medicamentos termolábiles.\n";
    if (tieneBajas) textoAnalisis += "• Temperatura BAJA (<2°C): riesgo de congelación de medicamentos refrigerados.\n";
    if (durTotal < 30) {
      textoAnalisis += "Los desvíos fueron breves. Se recomienda monitorear las próximas horas.";
    } else if (durTotal < 120) {
      textoAnalisis += "Desvíos de moderada duración. Evaluar posible afectación de medicamentos según Disposición ANMAT 10.872/2020.";
    } else {
      textoAnalisis += "Desvíos prolongados. Requiere evaluación técnica inmediata. Apartar medicamentos afectados con rótulo 'NO DISPENSAR - EN EVALUACIÓN' y notificar al Director Técnico de Farmacia.";
    }

    textoRecom = "";
    if (tieneAltas) {
      textoRecom += "• Temperatura ALTA: verificar sistema de refrigeración y sellado de puertas.\n";
      textoRecom += "  → Los medicamentos refrigerados pueden degradarse irreversiblemente por encima de 8°C.\n";
      textoRecom += "• Controlar termostato y estado del compresor.\n";
      textoRecom += "  → Un termostato descalibrado o compresor con falla son las causas más frecuentes de temperatura alta.\n";
      textoRecom += "• Evaluar la aptitud de los medicamentos afectados según protocolo vigente.\n";
      textoRecom += "  → La Disposición ANMAT 10.872/2020 exige evaluación documentada ante toda ruptura de cadena de frío.\n";
    }
    if (tieneBajas) {
      textoRecom += "• Temperatura BAJA: revisar configuración del termostato.\n";
      textoRecom += "  → La causa más frecuente es el termostato configurado demasiado frío.\n";
      textoRecom += "• Verificar que no haya contacto directo de medicamentos con el evaporador.\n";
      textoRecom += "• Controlar que la puerta no haya quedado abierta en ambiente frío.\n";
    }
    textoRecom += "• Documentar el evento en el registro de incidencias de cadena de frío.\n";
    textoRecom += "  → La normativa exige trazabilidad completa de toda ruptura para auditorías sanitarias.\n";
  }

  const notaTecnica = "NOTA TÉCNICA: Ante cualquier desvío térmico o falla del equipo, la intervención correctiva debe ser realizada por personal técnico calificado (Técnico en Refrigeración matriculado o servicio técnico autorizado por el fabricante). Ante fallas persistentes, contactar al Técnico en Refrigeración habilitado y notificar al Director Técnico de Farmacia según corresponda. Toda intervención debe quedar documentada con fecha, descripción y firma del responsable, conforme a la Disposición ANMAT 10.872/2020, la Disposición ANMAT 2069/2018 (Buenas Prácticas de Distribución) y las Buenas Prácticas de Almacenamiento (Resolución ANMAT 368/2000).";

  const notaResponsabilidad = "RESPONSABILIDAD: La responsabilidad del cumplimiento de las condiciones de conservación y de la cadena de frío en el sector Farmacia recae sobre el Director Técnico de Farmacia, conforme a la Ley Nacional 17.565 y su Decreto Reglamentario 7.123/68 (art. 10 inc. d), y la Disposición ANMAT 10.872/2020. Ante cualquier incidente, el Director Técnico debe ser notificado de forma inmediata.";

  return { alertasFilas, textoAnalisis, textoRecom, notaTecnica, notaResponsabilidad };
}

function generarGraficoCurva(feeds, field, nombre) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Tiempo")
    .addColumn(Charts.ColumnType.NUMBER, "°C");

  const vals = feeds.map(f => parseFloat(f[field])).filter(v => !isNaN(v) && v !== -127);
  if (vals.length === 0) {
    dataTable.addRow(["Sin datos", 5]);
    return Charts.newLineChart()
      .setDataTable(dataTable).setDimensions(2200, 520)
      .setColors(["#3b82f6"]).setOption("backgroundColor", "white")
      .build().getAs('image/png');
  }

  const minVal = Math.min(...vals);
  const maxVal = Math.max(...vals);
  const yMin = minVal - 0.5;
  const yMax = maxVal + 0.5;

  const UMBRAL_GAP_MS = 10 * 60 * 1000;
  const numPuntos = 800;
  const step = Math.max(1, Math.floor(feeds.length / numPuntos));

  const puntosMuestreados = [];
  for (let i = 0; i < feeds.length; i += step) puntosMuestreados.push(feeds[i]);

  let ultimoTsValido = null;

  for (let i = 0; i < puntosMuestreados.length; i++) {
    const f   = puntosMuestreados[i];
    const val = parseFloat(f[field]);
    const date = new Date(f.created_at);
    if (isNaN(date.getTime())) continue;

    // Detectar gap temporal — romper la línea solo por cortes reales de conectividad
    if (ultimoTsValido !== null && (date.getTime() - ultimoTsValido) > UMBRAL_GAP_MS) {
      dataTable.addRow([fmtFecha(date).slice(0, 13), null]);
    }

    // -127 o NaN: omitir el punto sin romper la línea
    if (!isNaN(val) && val !== -127) {
      dataTable.addRow([fmtFecha(date).slice(0, 13), val]);
      ultimoTsValido = date.getTime();
    }
  }

  return Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(2200, 520)
    .setColors(["#3b82f6"])
    .setOption("areaOpacity", 0.1)
    .setOption("lineWidth", 1.5)
    .setOption("vAxis", {
      gridlines: { count: 8, color: '#cbd5e1' },
      viewWindow: { min: yMin, max: yMax },
      format: '#.0°C',
      textStyle: { fontSize: 14, color: '#000000', bold: true },
      textPosition: 'out'
    })
    .setOption("hAxis", {
      slantedText: true, slantedTextAngle: 45,
      textStyle: { fontSize: 12, color: '#000000', bold: true },
      gridlines: { color: 'none' }, showTextEvery: 60
    })
    .setOption("chartArea", { width: '94%', height: '70%', left: '4%', right: '1%', top: '4%' })
    .setOption("legend", { position: 'none' })
    .setOption("backgroundColor", "white")
    .build().getAs('image/png');
}

function estilizarTabla(t) {
  const r0 = t.getRow(0);
  for(let i=0; i<r0.getNumCells(); i++) r0.getCell(i).setBackgroundColor("#f1f5f9").setBold(true).setFontSize(9);
  for(let i=1; i<t.getNumRows(); i++) {
    for(let j=0; j<t.getRow(i).getNumCells(); j++) t.getRow(i).getCell(j).setFontSize(8);
  }
}

function formatDur(m) {
  return m < 60 ? Math.round(m) + "m" : Math.floor(m/60) + "h " + Math.round(m%60) + "m";
}

function fetchThingSpeakData(id, key, d) {
  const res = UrlFetchApp.fetch(`https://api.thingspeak.com/channels/${id}/feeds.json?api_key=${key}&minutes=${d*1440}&results=8000`);
  return JSON.parse(res.getContentText());
}

/**
 * Obtiene TODOS los feeds de los últimos N días usando paginación hacia atrás.
 * Garantiza cubrir el período completo aunque haya más de 8000 registros.
 */
function fetchThingSpeakDataCompleto(id, key, dias) {
  const ahora  = new Date();
  const inicio = new Date(ahora.getTime() - dias * 24 * 60 * 60 * 1000);

  let todosLosFeeds = [];
  let fechaHasta    = new Date(ahora);
  let intentos      = 0;
  const MAX_INTENTOS = 10;

  while (intentos < MAX_INTENTOS) {
    const startStr = Utilities.formatDate(inicio,     "GMT-3", "yyyy-MM-dd'T'HH:mm:ss");
    const endStr   = Utilities.formatDate(fechaHasta, "GMT-3", "yyyy-MM-dd'T'HH:mm:ss");

    const url = `https://api.thingspeak.com/channels/${id}/feeds.json?api_key=${key}&start=${startStr}-03:00&end=${endStr}-03:00&results=8000`;
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(res.getContentText());

    if (!data.feeds || data.feeds.length === 0) break;

    const primerTs = todosLosFeeds.length > 0
      ? new Date(todosLosFeeds[0].created_at).getTime()
      : Infinity;
    const nuevos = data.feeds.filter(f => new Date(f.created_at).getTime() < primerTs);
    todosLosFeeds = nuevos.concat(todosLosFeeds);

    if (data.feeds.length < 8000) break;

    fechaHasta = new Date(new Date(data.feeds[0].created_at).getTime() - 60000);
    if (fechaHasta <= inicio) break;

    intentos++;
    Utilities.sleep(500);
  }

  const inicioMs = inicio.getTime();
  todosLosFeeds = todosLosFeeds.filter(f => new Date(f.created_at).getTime() >= inicioMs);
  return { feeds: todosLosFeeds };
}

function buscarLogoEnDrive(n) {
  const f = DriveApp.getFilesByName(n);
  return f.hasNext() ? f.next().getBlob() : null;
}

/**
 * Formatea una fecha a GMT-3 sin usar Utilities.formatDate (10-20x más rápido).
 */
function fmtFecha(date, conSegundos) {
  const d = new Date(date.getTime() - 3 * 60 * 60 * 1000);
  const p = n => String(n).padStart(2, '0');
  const base = `${p(d.getUTCDate())}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}`;
  return conSegundos ? base + ':' + p(d.getUTCSeconds()) : base;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'guardarPDF') {
      const bytes = Utilities.base64Decode(body.pdfData);
      const blob = Utilities.newBlob(bytes, 'application/pdf', body.filename);
      const file = DriveApp.getFolderById(CONFIG_FARMACIA.alertasFolder).createFile(blob);
      return ContentService.createTextOutput(JSON.stringify({result: true, url: file.getUrl()})).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({result: false, error: "Acción no reconocida"})).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({result: false, error: e.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ============================================================
 *  SHEET SEMANAL DE TEMPERATURAS
 *  Genera una pestaña nueva por semana en el Google Sheet
 *  configurado en CONFIG_FARMACIA.sheetId
 * ============================================================
 */
function generarSheetSemanal(feedsPorSensor, rangoTexto, fechaHoy) {
  const ss = SpreadsheetApp.openById(CONFIG_FARMACIA.sheetId);
  const nombreHoja = "Semana " + Utilities.formatDate(fechaHoy, "GMT-3", "dd-MM-yyyy");

  // Eliminar hoja si ya existe (re-ejecución)
  const hojaExistente = ss.getSheetByName(nombreHoja);
  if (hojaExistente) ss.deleteSheet(hojaExistente);

  const hoja = ss.insertSheet(nombreHoja);

  // ── ENCABEZADO PRINCIPAL ──────────────────────────────────
  hoja.getRange("A1").setValue("REGISTRO SEMANAL DE TEMPERATURAS - FARMACIA");
  hoja.getRange("A1").setFontSize(13).setFontWeight("bold").setFontColor("#00384d");
  hoja.getRange("A2").setValue("Período: " + rangoTexto);
  hoja.getRange("A2").setFontSize(10).setFontStyle("italic");
  hoja.getRange("A3").setValue("Generado: " + Utilities.formatDate(fechaHoy, "GMT-3", "dd/MM/yyyy HH:mm"));
  hoja.getRange("A3").setFontSize(9).setFontColor("#64748b");

  // ── CONSTRUIR COLUMNAS DINÁMICAMENTE ─────────────────────
  // Col 0: Fecha/Hora | Col 1..N: un sensor por columna
  const encabezados = ["Fecha / Hora"];
  feedsPorSensor.forEach(fs => {
    encabezados.push(fs.sensor.n + "\n(" + fs.sensor.eq + ")");
  });

  const filaEncabezado = 5;
  const rangoEnc = hoja.getRange(filaEncabezado, 1, 1, encabezados.length);
  rangoEnc.setValues([encabezados]);
  rangoEnc.setBackground("#00384d").setFontColor("#ffffff").setFontWeight("bold")
          .setFontSize(10).setWrap(true).setVerticalAlignment("middle")
          .setHorizontalAlignment("center");
  hoja.setRowHeight(filaEncabezado, 45);

  // ── UNIFICAR TIMESTAMPS AGRUPANDO POR MINUTO ─────────────
  const mapaTemp = {};
  feedsPorSensor.forEach((fs, idx) => {
    fs.feeds.forEach(feed => {
      const val = parseFloat(feed[fs.sensor.field]);
      if (isNaN(val) || val === -127) return;
      const clave = fmtFecha(new Date(feed.created_at));
      if (!mapaTemp[clave]) mapaTemp[clave] = { fecha: clave, valores: {} };
      if (mapaTemp[clave].valores[idx] !== undefined) {
        mapaTemp[clave].valores[idx] = (mapaTemp[clave].valores[idx] + val) / 2;
      } else {
        mapaTemp[clave].valores[idx] = val;
      }
    });
  });

  // Ordenar por clave de fecha (formato dd/MM/yyyy HH:mm → convertir para ordenar)
  const claves = Object.keys(mapaTemp).sort((a, b) => {
    const toDate = s => {
      const [fecha, hora] = s.split(' ');
      const [d, m, y] = fecha.split('/');
      return new Date(`${y}-${m}-${d}T${hora}:00`);
    };
    return toDate(a) - toDate(b);
  });

  // ── ESCRIBIR DATOS EN LOTES (más eficiente) ───────────────
  const filas = claves.map(clave => {
    const entrada = mapaTemp[clave];
    const fila = [entrada.fecha];
    feedsPorSensor.forEach((_, idx) => {
      const v = entrada.valores[idx];
      fila.push(v !== undefined ? Math.round(v * 100) / 100 : "");
    });
    return fila;
  });

  if (filas.length > 0) {
    const filaInicio = filaEncabezado + 1;
    hoja.getRange(filaInicio, 1, filas.length, encabezados.length).setValues(filas);

    // ── FORMATO CONDICIONAL en batch ──────────────────────────
    const todasLasReglas = [];
    feedsPorSensor.forEach((_, idx) => {
      const col = idx + 2;
      const rangoCol = hoja.getRange(filaInicio, col, filas.length, 1);
      todasLasReglas.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(8.0)
          .setBackground("#fecaca").setFontColor("#dc2626")
          .setRanges([rangoCol]).build(),
        SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(2.0)
          .setBackground("#bfdbfe").setFontColor("#1d4ed8")
          .setRanges([rangoCol]).build()
      );
    });
    hoja.setConditionalFormatRules(todasLasReglas);

    // ── FORMATO DE COLUMNAS ───────────────────────────────────
    hoja.setColumnWidth(1, 140);
    feedsPorSensor.forEach((_, idx) => hoja.setColumnWidth(idx + 2, 130));

    // Centrar columnas de temperatura
    hoja.getRange(filaInicio, 2, filas.length, feedsPorSensor.length)
        .setHorizontalAlignment("center").setNumberFormat("0.00");
  }

  // ── FILA DE RESUMEN ESTADÍSTICO en batch ──────────────────
  const filaResumen = filaEncabezado + filas.length + 2;
  hoja.getRange(filaResumen, 1).setValue("RESUMEN ESTADÍSTICO")
      .setFontWeight("bold").setFontColor("#00384d").setFontSize(10);

  const etiquetas = ["Mínimo (°C)", "Máximo (°C)", "Promedio (°C)", "Lecturas totales"];
  hoja.getRange(filaResumen + 1, 1, etiquetas.length, 1)
      .setValues(etiquetas.map(e => [e])).setFontWeight("bold");

  feedsPorSensor.forEach((fs, idx) => {
    const col = idx + 2;
    const valores = filas
      .map(f => f[col - 1])
      .filter(v => v !== "" && !isNaN(v))
      .map(Number);

    let stats;
    if (valores.length > 0) {
      const minVal = Math.min(...valores);
      const maxVal = Math.max(...valores);
      const avg = Math.round((valores.reduce((a, b) => a + b, 0) / valores.length) * 100) / 100;
      stats = [[minVal], [maxVal], [avg], [valores.length]];
    } else {
      stats = [["--"], ["--"], ["--"], [0]];
    }
    hoja.getRange(filaResumen + 1, col, 4, 1).setValues(stats);
  });

  // Estilo del bloque resumen
  hoja.getRange(filaResumen + 1, 1, 4, encabezados.length)
      .setBackground("#f1f5f9").setBorder(true, true, true, true, true, true);

  // Congelar fila de encabezado
  hoja.setFrozenRows(filaEncabezado);

  console.log("Sheet semanal generado: " + nombreHoja + " (" + filas.length + " registros)");
}

/** Convierte número de columna a letra (1→A, 2→B, 27→AA, etc.) */
function columnToLetter(col) {
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Función de prueba — ejecutar desde el editor de Apps Script
 * para generar el Sheet semanal sin esperar el trigger automático.
 */
function probarSheetSemanal() {
  const hoy = new Date();
  const haceSieteDias = new Date(hoy.getTime() - 7 * 24 * 60 * 60 * 1000);
  const rangoTexto = Utilities.formatDate(haceSieteDias, "GMT-3", "dd/MM/yyyy") + 
                     " - " + 
                     Utilities.formatDate(hoy, "GMT-3", "dd/MM/yyyy");

  const feedsPorSensor = [];
  SENSORES.forEach(s => {
    try {
      const data = fetchThingSpeakDataCompleto(s.id, s.k, 7);
      if (!data || !data.feeds || data.feeds.length === 0) return;
      feedsPorSensor.push({ sensor: s, feeds: data.feeds });
    } catch (e) {
      console.error("Error cargando sensor " + s.n + ": " + e.message);
    }
  });

  if (feedsPorSensor.length === 0) {
    console.error("No se obtuvieron datos de ningún sensor.");
    return;
  }

  generarSheetSemanal(feedsPorSensor, rangoTexto, hoy);
}
