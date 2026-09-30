// Google Apps Script: Extensiones > Apps Script en tu Google Sheet.
// Implementar > Nueva implementación > Aplicación web > Acceso: "Cualquier persona".
// Copia la URL /exec en SHEETS_URL de index.html.
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  if (d.tipo === 'rsvp') {
    sheet(ss, 'Asistencia', ['Fecha','Nombre','Asistentes','Mensaje'])
      .appendRow([now, d.nombre, d.asistentes, d.mensaje]);
  } else {
    sheet(ss, 'Mensajes', ['Fecha','Nombre','Para','Privado','Texto','Foto','Audio'])
      .appendRow([now, d.nombre, d.para, d.privado, d.texto, save(d.foto, 'foto_' + now.getTime() + '.jpg'), save(d.audio, 'voz_' + now.getTime() + '.webm')]);
  }
  return ContentService.createTextOutput('ok');
}
function sheet(ss, name, head) {
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); }
  return s;
}
function save(dataUrl, name) {
  if (!dataUrl) return '';
  const m = dataUrl.match(/^data:(.*?)(;.*)?;base64,(.*)$/);
  if (!m) return '';
  const it = DriveApp.getFoldersByName('Recuerdos boda');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Recuerdos boda');
  return folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], name)).getUrl();
}
