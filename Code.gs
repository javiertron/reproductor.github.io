// Google Apps Script — recibe fotos subidas desde upload.html y las guarda en Drive.
// 1. Ve a https://script.google.com > Nuevo proyecto, pega este código.
// 2. Cambia FOLDER_ID por el ID de la carpeta de Drive donde quieres guardar las fotos
//    (créala primero, ábrela, y copia el ID que aparece en la URL después de /folders/).
// 3. Implementar > Nueva implementación > Aplicación web.
//    - Ejecutar como: Yo
//    - Quién tiene acceso: Cualquier usuario
// 4. Copia la URL que te da y pégala en SCRIPT_URL dentro de upload.html.
// 5. Opcional: crea una Hoja de cálculo y pon su ID en SHEET_ID para llevar un registro
//    con nombre, mensaje y hora de cada envío (déjalo vacío para omitir este paso).

const FOLDER_ID = "PASTE_YOUR_DRIVE_FOLDER_ID_HERE";
const SHEET_ID = ""; // opcional

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const savedNames = [];

    (body.files || []).forEach(f => {
      const bytes = Utilities.base64Decode(f.data);
      const blob = Utilities.newBlob(bytes, f.mimeType, f.filename);
      const file = folder.createFile(blob);
      savedNames.push(file.getName());
    });

    if (SHEET_ID) {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
      sheet.appendRow([
        new Date(),
        body.name || "",
        body.message || "",
        savedNames.join(", ")
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, saved: savedNames.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
