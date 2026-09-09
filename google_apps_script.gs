const SHEET_NAME = "Confirmaciones";

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = {};

    try {
      data = JSON.parse(raw);
    } catch (_) {
      // Respaldo por si alguna versión futura envía campos de formulario.
      data = (e && e.parameter) ? e.parameter : {};
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Mantener los encabezados del diseño actual en la fila 6.
    if (sheet.getRange("A6").getValue() !== "FECHA Y HORA") {
      sheet.getRange("A6:E6").setValues([[
        "FECHA Y HORA", "INVITADO", "EVENTO", "RESPUESTA", "PERSONAS"
      ]]);
    }

    var nextRow = Math.max(sheet.getLastRow() + 1, 7);

    sheet.getRange(nextRow, 1, 1, 5).setValues([[
      new Date(),
      String(data.nombre || ""),
      String(data.evento || ""),
      String(data.respuesta || ""),
      Number(data.personas || 0)
    ]]);

    // Formatos importantes: resumen numérico y fechas solo en registros.
    sheet.getRange("A4:E4").setNumberFormat("0");
    sheet.getRange("A7:A").setNumberFormat("dd/MM/yyyy HH:mm");

    return ContentService
      .createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ok:true,service:"RSVP Ethan"}))
    .setMimeType(ContentService.MimeType.JSON);
}
