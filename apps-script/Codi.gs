// Canvieu la contrasenya abans de publicar
const CLAU = 'CANVIEU-ME';
const FULL = 'Respostes';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FULL);
  if (!sh) {
    sh = ss.insertSheet(FULL);
    sh.appendRow(['Data', 'Nom', 'Assisteix', 'Persones', 'Al·lèrgies', 'Missatge']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.web) return ContentService.createTextOutput('ok');
    const clean = s => String(s || '').slice(0, 1000).replace(/^[=+\-@]/, "'$&");
    sheet_().appendRow([
      new Date(), clean(d.nom), d.assistencia === 'Sí' ? 'Sí' : 'No',
      Math.max(0, Math.min(10, parseInt(d.persones, 10) || 0)), clean(d.alergies), clean(d.missatge)
    ]);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!e.parameter.clau || e.parameter.clau !== CLAU) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false })).setMimeType(ContentService.MimeType.JSON);
  }
  const v = sheet_().getDataRange().getValues().slice(1);
  const tz = Session.getScriptTimeZone();
  const rows = v.map(r => ({
    data: Utilities.formatDate(new Date(r[0]), tz, 'dd/MM/yyyy HH:mm'),
    nom: r[1], assistencia: r[2], persones: r[3], alergies: r[4], missatge: r[5]
  })).reverse();
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows })).setMimeType(ContentService.MimeType.JSON);
}
