// Google Apps Script — Diagnóstico de Eficiencia Operacional
// Instrucciones:
// 1. Ve a tu Google Sheet → Extensiones → Apps Script
// 2. Borra el contenido existente y pega este código completo
// 3. Guarda (Ctrl+S)
// 4. Haz clic en "Implementar" → "Nueva implementación" (o "Gestionar implementaciones" si ya existe)
// 5. Tipo: "Aplicación web" | Ejecutar como: "Yo" | Acceso: "Cualquier persona"
// 6. Copia la URL y pégala en el HTML (variable APPS_SCRIPT_URL)

const SHEET_NAME = 'Respuestas';

function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    const data  = JSON.parse(e.postData.contents);

    if (sheet.getLastRow() === 0) {
      const headers = sheet.appendRow(getHeaders());
      sheet.getRange(1, 1, 1, getHeaders().length)
        .setFontWeight('bold')
        .setBackground('#1a1a1a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow(buildRow(data));
    sheet.autoResizeColumns(1, getHeaders().length);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'Endpoint activo' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function getHeaders() {
  return [
    'Fecha y hora',
    // Datos de contacto
    'Nombre',
    'Empresa',
    'Cargo',
    'Email',
    'Teléfono',
    // Respuestas del diagnóstico
    'Personas en el equipo',
    'Costo promedio mensual',
    'Tareas repetitivas',
    'Automatización actual',
    'Horas de reunión / semana',
    'Seguimiento de compromisos',
    'Frecuencia de retrabajo',
    'Uso de IA / herramientas',
    // Resultados calculados
    'Pérdida total estimada (anual)',
    'Pérdida como % del costo',
    'Recuperable en 90 días',
    'Nivel de madurez'
  ];
}

function buildRow(d) {
  const costoMap = { a:'Menos de $800.000', b:'$800.000 – $1.500.000', c:'$1.500.000 – $2.500.000', d:'Más de $2.500.000' };
  const repMap   = { a:'Menos del 10%', b:'10% – 25%', c:'25% – 40%', d:'Más del 40%' };
  const autoMap  = { a:'Ninguna — todo manual', b:'Algunas, de forma parcial', c:'La mayoría automatizadas' };
  const seguMap  = { a:'Hay seguimiento formal', b:'Seguimiento irregular', c:'Depende de cada persona', d:'Generalmente no se registran' };
  const retraMap = { a:'Casi nunca', b:'Una o dos veces al mes', c:'Varias veces por semana', d:'Es parte del día a día' };
  const iaMap    = { a:'No usamos ninguna', b:'Experimentamos sin sistematicidad', c:'Sí, uso regular' };

  return [
    new Date(),
    d.nombre      || '',
    d.empresa     || '',
    d.cargo       || '',
    d.email       || '',
    d.telefono    || '',
    d.personas,
    costoMap[d.o2]  || d.o2,
    repMap[d.o3]    || d.o3,
    autoMap[d.o4]   || d.o4,
    d.horasReunion + 'h / semana',
    seguMap[d.o6]   || d.o6,
    retraMap[d.o7]  || d.o7,
    iaMap[d.o8]     || d.o8,
    d.perdidaTotal,
    d.pctTotal + '%',
    d.recuperable90,
    d.nivelMadurez
  ];
}
