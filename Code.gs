/**
 * UREDINTERACTIVA2026
 * Base de datos central de evaluaciones de RED.
 *
 * 1. Cree una Hoja de cálculo de Google.
 * 2. Extensiones > Apps Script.
 * 3. Cree un archivo HTML llamado Index y pegue el contenido de Index.html.
 * 4. Pegue este código en Code.gs.
 * 5. Ejecute configurarHoja() una vez y autorice los permisos.
 * 6. Implementar > Nueva implementación > Aplicación web.
 *    Ejecutar como: Yo
 *    Quién tiene acceso: Cualquier persona
 * 7. Comparta la URL que genera Google.
 */

const NOMBRE_HOJA = "Evaluaciones";

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("UREDINTERACTIVA2026 - Evaluación de RED")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function configurarHoja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOMBRE_HOJA);
  if (!sh) sh = ss.insertSheet(NOMBRE_HOJA);

  const encabezados = [
    "Fecha","Nombre del RED","Evaluador",
    "Diseño y presentación","Usabilidad","Motivación","Funcionalidad",
    "Rendimiento","Soportabilidad","Confiabilidad",
    "Puntaje","Promedio","Porcentaje","Nivel","Observaciones","Fecha de registro"
  ];
  sh.clear();
  sh.getRange(1,1,1,encabezados.length).setValues([encabezados]);
  sh.getRange(1,1,1,encabezados.length).setFontWeight("bold");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1,encabezados.length);
  return "Hoja configurada correctamente.";
}

function guardarEvaluacion(data) {
  if (!data || !data.red || !data.evaluador) {
    throw new Error("Faltan datos obligatorios.");
  }
  if (!Array.isArray(data.scores) || data.scores.length !== 7 || data.scores.some(x => Number(x) < 1 || Number(x) > 5)) {
    throw new Error("Debe seleccionar una valoración de 1 a 5 para los 7 factores.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(NOMBRE_HOJA);
  if (!sh) {
    configurarHoja();
    sh = ss.getSheetByName(NOMBRE_HOJA);
  }

  sh.appendRow([
    data.fecha || new Date(),
    data.red,
    data.evaluador,
    ...data.scores,
    Number(data.total),
    Number(data.promedio),
    Number(data.porcentaje),
    data.nivel,
    data.observaciones || "",
    new Date()
  ]);

  return {ok:true, message:"¡Evaluación guardada correctamente en la base de datos central!"};
}
