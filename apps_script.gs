// Apps Script v2. Reemplaza TODO el código y luego: Implementar > Administrar implementaciones > lápiz > Nueva versión.
// Antes, borra la pestaña "Mensajes" de prueba para que se cree con las columnas nuevas.
const HEAD = ['Fecha','ID','Nombre','Para','Privado','Texto','Foto','Audio','FotoId','AudioId','Publicar'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();
    if (d.tipo === 'rsvp') {
      const s = sheet(ss, 'Asistencia', ['Fecha','Nombre','Asistentes','Mensaje']);
      const rows = s.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (norm(rows[i][1]) === norm(d.nombre)) {   // misma persona: se actualiza su respuesta
          s.getRange(i + 1, 1, 1, 4).setValues([[now, t(d.nombre), d.asistentes, t(d.mensaje)]]);
          return out({ ok: true, updated: true });
        }
      }
      s.appendRow([now, t(d.nombre), d.asistentes, t(d.mensaje)]);
      return out({ ok: true, updated: false });
    }
    const s = sheet(ss, 'Mensajes', HEAD);
    const id = Utilities.getUuid().slice(0, 8);
    const foto = save(d.foto, 'foto_' + id + '.jpg');
    const audio = save(d.audio, 'voz_' + id + '.' + ((d.audio || '').indexOf('mp4') > -1 ? 'm4a' : 'webm'));
    s.appendRow([now, id, t(d.nombre), d.para, d.privado === true, t(d.texto), foto.url, audio.url, foto.id, audio.id, false]);
    s.getRange(s.getLastRow(), 11).insertCheckboxes();   // marca "Publicar" para mostrarlo en el muro
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Solo entrega lo marcado en "Publicar" y que no sea privado.
function doGet(e) {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Mensajes');
  if (!s) return out({ items: [] });
  const rows = s.getDataRange().getValues().slice(1).filter(r => r[10] === true && r[4] !== true);
  const p = (e && e.parameter) || {};
  if (p.media) {
    const r = rows.find(r => r[8] === p.media || r[9] === p.media);
    if (!r) return out({ error: 'no disponible' });
    const f = DriveApp.getFileById(p.media);
    return out({ data: 'data:' + f.getMimeType() + ';base64,' + Utilities.base64Encode(f.getBlob().getBytes()) });
  }
  return out({ items: rows.map(r => ({
    id: r[1], nombre: r[2], para: r[3], texto: r[5], foto: r[8] || '', audio: r[9] || '',
    fecha: Utilities.formatDate(new Date(r[0]), 'America/Mexico_City', 'dd/MM/yyyy')
  })).reverse() });
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim(); }
function t(v) { v = String(v || ''); return /^[=+\-@]/.test(v) ? ' ' + v : v; }   // evita que Sheets lo lea como fórmula
function sheet(ss, name, head) {
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); }
  return s;
}
function save(dataUrl, name) {
  if (!dataUrl) return { url: '', id: '' };
  const m = dataUrl.match(/^data:(.*?)(;.*)?;base64,(.*)$/);
  if (!m) return { url: '', id: '' };
  const it = DriveApp.getFoldersByName('Recuerdos boda');
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Recuerdos boda');
  const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[3]), m[1], name));
  return { url: f.getUrl(), id: f.getId() };
}
