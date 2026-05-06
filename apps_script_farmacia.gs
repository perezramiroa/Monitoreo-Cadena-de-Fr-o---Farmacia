/**
 * ============================================================
 *  VICUS - FARMACIA  |  Google Apps Script (PROFESIONAL V4)
 *  Hospital Natalio Burd
 *  Desarrollado por Ingeniero Perez Ramiro
 * ============================================================
 */

const CONFIG_FARMACIA = {
  folderPDF: '1Kd-8NUFiWCVuiu4enxDUDdgfX0UmFLvu',
  sheetId: '1g4SLfsHrFfhsbzuY-WnTzl3gmTgyUbkx',
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
  
  SENSORES.forEach(s => {
    try {
      const data = fetchThingSpeakData(s.id, s.k, 7);
      if (!data || !data.feeds || data.feeds.length === 0) return;

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
}

function generarPDFOficial(sensor, fecha, rango, trazabilidad, analizada, conectividad, grafico) {
  const doc = DocumentApp.create('Temp_Reporte_' + sensor.n);
  const body = doc.getBody();

  // Configurar márgenes al mínimo posible para maximizar las tablas y el gráfico
  body.setMarginLeft(20).setMarginRight(20).setMarginTop(20).setMarginBottom(20);
  const anchoMax = 555; // 595 - 40

  // --- CABECERA ---
  const logo = buscarLogoEnDrive("logo_rih.jpg");
  if (logo) {
    const header = doc.addHeader();
    const hp = header.appendParagraph("");
    hp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    hp.appendInlineImage(logo).setWidth(anchoMax).setHeight(60); // Reducido de 85 a 60
  }

  const t1 = body.appendParagraph("INFORME TÉCNICO DE CADENA DE FRÍO\nMEDICAMENTOS REFRIGERADOS");
  t1.setFontSize(14).setBold(true).setForegroundColor("#00384d");
  body.appendParagraph("Según Disposición ANMAT 10.872/2020").setFontSize(9).setItalic(true);

  body.appendParagraph("\nDispositivo: " + sensor.n).setBold(true).setFontSize(11);
  body.appendParagraph("Equipo/Artefacto: " + sensor.eq).setItalic(true).setFontSize(10);
  body.appendParagraph("Período: " + rango).setBold(true).setFontSize(10);
  body.appendParagraph("Emisión: " + fecha + " | Trazabilidad: " + trazabilidad).setFontSize(8);
  
  // Línea separadora después del encabezado
  const separador1 = body.appendParagraph("");
  separador1.setBorder(DocumentApp.ParagraphHeading.NORMAL, DocumentApp.BorderPosition.BOTTOM, 1, "#cbd5e1", 0);
  separador1.setSpacingAfter(8);

  // --- GRÁFICO (ANCHO COMPLETO IGUAL A LAS TABLAS) ---
  body.appendParagraph("\nCURVA TÉRMICA SEMANAL").setBold(true).setFontSize(10);
  const pChart = body.appendParagraph("");
  pChart.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  pChart.appendInlineImage(grafico).setWidth(anchoMax).setHeight(240); // Reducido altura de 260 a 240

  // --- TABLA DE ALERTAS ---
  body.appendParagraph("\n⚠️ ALERTAS Y RECUPERACIONES (2°C - 8°C)").setBold(true).setFontSize(10);
  const tablaAlertas = [["Fecha y Hora", "Valor", "Estado", "Duración"]];
  if (analizada.alertasFilas.length > 0) {
    analizada.alertasFilas.forEach(f => tablaAlertas.push([f.h, f.v, f.e, f.d]));
  } else {
    tablaAlertas.push(["-", "-", "✅ Sin eventos fuera de rango", "-"]);
  }
  estilizarTabla(body.appendTable(tablaAlertas));

  // --- EVENTOS DE CONECTIVIDAD (SIN COLUMNA OBS.) ---
  body.appendParagraph("\n📡 EVENTOS DETECTADOS (>10 min sin datos)").setBold(true).setFontSize(10);
  const tablaWifi = [["Inicio", "Fin", "Tipo de Corte", "T. Antes", "T. Desp.", "Duración"]]; // Eliminada columna Obs.
  if (conectividad.filas.length > 0) {
    conectividad.filas.forEach(f => tablaWifi.push([f.inicio, f.fin, f.tipo, f.antes, f.despues, f.duracion])); // Sin ""
  } else {
    tablaWifi.push(["-", "-", "✅ Sin interrupciones significativas", "-", "-", "-"]); // Sin columna extra
  }
  estilizarTabla(body.appendTable(tablaWifi));

  // --- ANÁLISIS Y RECOMENDACIONES (SEPARADOS VISUALMENTE) ---
  body.appendParagraph("\nANÁLISIS TÉCNICO:").setBold(true).setFontSize(10);
  
  // Análisis de alertas
  if (analizada.textoAnalisis) {
    body.appendParagraph(analizada.textoAnalisis).setFontSize(9).setItalic(true);
  }
  
  // Análisis de conectividad (separado)
  if (conectividad.analisis && conectividad.analisis !== "Sin problemas de conectividad.") {
    body.appendParagraph("\nConectividad:").setBold(true).setFontSize(9);
    body.appendParagraph(conectividad.analisis).setFontSize(9).setItalic(true);
  }

  body.appendParagraph("\nRECOMENDACIONES:").setBold(true).setFontSize(10).setForegroundColor("#00384d");
  body.appendParagraph(analizada.textoRecom + "\n• " + conectividad.recom).setFontSize(9);
  
  // Línea separadora antes del footer
  const separador2 = body.appendParagraph("");
  separador2.setBorder(DocumentApp.ParagraphHeading.NORMAL, DocumentApp.BorderPosition.TOP, 1, "#cbd5e1", 0);
  separador2.setSpacingBefore(12);

  // --- PIE DE PÁGINA ---
  const footer = doc.addFooter();
  const logoF = buscarLogoEnDrive("footer.jpg");
  if (logoF) {
    const fp = footer.appendParagraph("");
    fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    fp.appendInlineImage(logoF).setWidth(anchoMax).setHeight(50); // Reducido de 65 a 50
  }

  doc.saveAndClose();
  const pdf = doc.getAs('application/pdf');
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  return pdf;
}

function analizarConectividad(feeds, field) {
  let filas = [];
  let totalMinutos = 0;
  for (let i = 1; i < feeds.length; i++) {
    const d1 = new Date(feeds[i-1].created_at);
    const d2 = new Date(feeds[i].created_at);
    const diff = (d2 - d1) / 60000;
    if (diff > 10) {
      const v1 = parseFloat(feeds[i-1][field]);
      const v2 = parseFloat(feeds[i][field]);
      const tipo = (v2 > 8 || v2 < 2) ? 'Corte Energía' : 'Corte WiFi';
      filas.push({
        inicio: Utilities.formatDate(d1, "GMT-3", "dd/MM HH:mm"),
        fin: Utilities.formatDate(d2, "GMT-3", "dd/MM HH:mm"),
        tipo: tipo,
        antes: isNaN(v1) ? "--" : v1.toFixed(2) + "°C",
        despues: isNaN(v2) ? "--" : v2.toFixed(2) + "°C",
        duracion: formatDur(diff)
      });
      totalMinutos += diff;
    }
  }
  return {
    filas: filas,
    analisis: filas.length > 0 ? `Se detectaron ${filas.length} ${filas.length === 1 ? 'interrupción' : 'interrupciones'} de datos.\n• Tiempo total sin datos: ${formatDur(totalMinutos)}` : "Sin problemas de conectividad.",
    recom: "Verificar el estado del router y la conexión a internet.\n• Revisar la distancia entre el sensor y el punto de acceso WiFi."
  };
}

function analizarDatos(feeds, field) {
  let alertasFilas = [];
  let lastState = 'normal';
  let startTime = null;
  let stats = [];
  feeds.forEach(f => {
    const val = parseFloat(f[field]);
    if (isNaN(val)) return;
    const state = (val > 8.0) ? 'Alta' : (val < 2.0) ? 'Baja' : 'normal';
    if (state !== lastState) {
      const hora = Utilities.formatDate(new Date(f.created_at), "GMT-3", "dd/MM HH:mm");
      if (state !== 'normal') {
        startTime = new Date(f.created_at);
        alertasFilas.push({ h: hora, v: val.toFixed(1) + "°C", e: "⚠️ Alerta " + state, d: "--" });
      } else if (startTime) {
        const dur = (new Date(f.created_at) - startTime) / 60000;
        alertasFilas.push({ h: hora, v: val.toFixed(1) + "°C", e: "✅ Recuperación", d: formatDur(dur) });
        stats.push({ s: lastState, d: dur });
      }
      lastState = state;
    }
  });
  return {
    alertasFilas: alertasFilas,
    textoAnalisis: stats.length > 0 ? `Se detectaron ${stats.length} desvíos térmicos (${formatDur(stats.reduce((a,b)=>a+b.d,0))}).` : "Estabilidad térmica confirmada.",
    textoRecom: stats.length > 0 ? "• Revisar sellado de puertas.\n• Controlar termostato." : "• Continuar monitoreo habitual."
  };
}

function generarGraficoCurva(feeds, field, nombre) {
  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, "Tiempo")
    .addColumn(Charts.ColumnType.NUMBER, "°C");

  let vals = feeds.map(f => parseFloat(f[field])).filter(v => !isNaN(v));
  if (vals.length === 0) vals = [5];
  let minVal = Math.min(...vals);
  let maxVal = Math.max(...vals);

  let yMin = minVal - 0.5;
  let yMax = maxVal + 0.5;

  const numPuntos = 800; // Mucho más detalle para igualar al reporte manual
  const step = Math.max(1, Math.floor(feeds.length / numPuntos));
  
  for (let i = 0; i < feeds.length; i += step) {
    let f = feeds[i];
    let val = parseFloat(f[field]);
    let date = new Date(f.created_at);
    
    if (!isNaN(val) && !isNaN(date.getTime())) {
      let label = Utilities.formatDate(date, "GMT-3", "dd/MM HH:mm");
      dataTable.addRow([label, val]);
    }
  }

  return Charts.newLineChart()
    .setDataTable(dataTable)
    .setDimensions(1400, 550) // Más alto para leyendas grandes
    .setColors(["#3b82f6"]) 
    .setOption("areaOpacity", 0.1) 
    .setOption("lineWidth", 1.5) 
    .setOption("vAxis", { 
      gridlines: { count: 8, color: '#cbd5e1' }, 
      viewWindow: { min: yMin, max: yMax },
      format: '#.0°C',
      textStyle: { fontSize: 18, color: '#000000', bold: true },
      textPosition: 'in'
    })
    .setOption("hAxis", { 
      slantedText: true, 
      slantedTextAngle: 45,
      textStyle: { fontSize: 14, color: '#000000', bold: true }, 
      gridlines: { color: 'none' },
      showTextEvery: 60 
    })
    .setOption("chartArea", { width: '98%', height: '65%', left: '0%', right: '2%', top: '5%' })
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
  const res = UrlFetchApp.fetch(`https://api.thingspeak.com/channels/${id}/feeds.json?api_key=${key}&minutes=${d*1440}`);
  return JSON.parse(res.getContentText());
}

function buscarLogoEnDrive(n) {
  const f = DriveApp.getFilesByName(n);
  return f.hasNext() ? f.next().getBlob() : null;
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  if (body.action === 'guardarPDF') {
    const bytes = Utilities.base64Decode(body.pdfData);
    const blob = Utilities.newBlob(bytes, 'application/pdf', body.filename);
    const file = DriveApp.getFolderById(CONFIG_FARMACIA.alertasFolder).createFile(blob);
    return ContentService.createTextOutput(JSON.stringify({result: true, url: file.getUrl()})).setMimeType(ContentService.MimeType.JSON);
  }
}
