// Canvieu la contrasenya abans de publicar
const CLAU = 'CANVIEU-ME';
const FULL = 'Convidats';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FULL);
  if (!sh) {
    sh = ss.insertSheet(FULL);
    sh.appendRow(['Data', 'Confirmat per', 'Nom i cognoms', 'Assisteix', 'Al·lèrgies / intoleràncies', 'Missatge']);
    sh.setFrozenRows(1);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#e9eee3');
    sh.setColumnWidths(1, 1, 130); sh.setColumnWidths(2, 2, 200); sh.setColumnWidths(4, 1, 80); sh.setColumnWidths(5, 2, 260);
  }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.web) return ContentService.createTextOutput('ok');
    const clean = s => String(s || '').slice(0, 300).replace(/^[=+\-@]/, "'$&");
    const assisteix = d.assistencia === 'Sí' ? 'Sí' : 'No';
    const llista = (Array.isArray(d.convidats) ? d.convidats : []).slice(0, 10);
    const now = new Date();
    const rows = llista.map((c, i) => [now, clean(d.nom), clean(c.nom), assisteix, clean(c.alergies), i === 0 ? clean(d.missatge) : '']);
    if (rows.length) sheet_().getRange(sheet_().getLastRow() + 1, 1, rows.length, 6).setValues(rows);
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
    grup: r[1], nom: r[2], assistencia: r[3], alergies: r[4], missatge: r[5]
  })).reverse();
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows })).setMimeType(ContentService.MimeType.JSON);
}
