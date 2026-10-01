// Pravin & Saranya wedding — backend (Google Apps Script)
//  · guest photo album (Drive folder)      · shared sky of wish lanterns (auto-created Sheet)
//  · live heart counter (Script Properties)
// Deploy: Deploy > Manage deployments > Edit > Version: New version > Deploy
// First time after pasting: pick function "setup" in the toolbar, press Run, and Allow access.
var FOLDER_ID = '1Gac2OpZJZBdPxEUolfGXSy7r6qHLeg96';
var MAX_BYTES = 6 * 1024 * 1024;
var MAX_LIST  = 60;
var MAX_LANTERNS = 80;

function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function props_(){return PropertiesService.getScriptProperties();}

// Creates the lantern sheet once and remembers it
function sheet_(){
  var p = props_(), id = p.getProperty('SKY_SHEET'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create('Wedding sky: lanterns (delete a row to remove a wish)');
    p.setProperty('SKY_SHEET', ss.getId());
    var sh = ss.getSheets()[0]; sh.setName('lanterns'); sh.appendRow(['time', 'name', 'wish']);
  }
  return ss.getSheetByName('lanterns') || ss.getSheets()[0];
}
function setup(){ sheet_(); DriveApp.getFolderById(FOLDER_ID); Logger.log('Ready: ' + SpreadsheetApp.openById(props_().getProperty('SKY_SHEET')).getUrl()); }

function lanterns_(){
  var sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  var start = Math.max(2, n - MAX_LANTERNS + 1), rows = sh.getRange(start, 1, n - start + 1, 3).getValues();
  return rows.map(function(r, i){ return {id: start + i, n: String(r[1]), w: String(r[2])}; });
}
function clean_(t, max){
  t = String(t || '').replace(/[\u0000-\u001f<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;   // no spreadsheet formulas
}

function doGet(e){
  var a = e && e.parameter && e.parameter.a;
  if (a === 'sky') return json_({hearts: Number(props_().getProperty('HEARTS') || 0), lanterns: lanterns_()});
  if (a === 'heart') {
    var n = Math.max(0, Math.min(30, parseInt(e.parameter.n, 10) || 0)), lock = LockService.getScriptLock(), c;
    lock.waitLock(5000);
    try { c = Number(props_().getProperty('HEARTS') || 0) + n; props_().setProperty('HEARTS', String(c)); } finally { lock.releaseLock(); }
    return json_({hearts: c});
  }
  // default: newest photos (ids only; the page builds thumbnail URLs)
  var it = DriveApp.getFolderById(FOLDER_ID).getFiles(), out = [];
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType().indexOf('image/') === 0) {
      if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      out.push({id: f.getId(), t: f.getDateCreated().getTime()});
    }
  }
  out.sort(function(a, b){return b.t - a.t;});
  return json_(out.slice(0, MAX_LIST));
}

function doPost(e){
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.kind === 'lantern') {
      var wish = clean_(d.wish, 80), name = clean_(d.name, 30);
      if (!wish) throw new Error('empty');
      var lock = LockService.getScriptLock(); lock.waitLock(5000);
      try { var sh = sheet_(); sh.appendRow([new Date(), name, wish]); return json_({ok: true, id: sh.getLastRow()}); }
      finally { lock.releaseLock(); }
    }
    if (!/^image\/(jpeg|png|webp)$/.test(d.type)) throw new Error('type');
    var bytes = Utilities.base64Decode(d.data);
    if (bytes.length > MAX_BYTES) throw new Error('size');
    var ext = d.type === 'image/png' ? 'png' : d.type === 'image/webp' ? 'webp' : 'jpg';
    var file = DriveApp.getFolderById(FOLDER_ID).createFile(Utilities.newBlob(bytes, d.type, 'guest-' + Date.now() + '.' + ext));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return json_({ok: true, id: file.getId()});
  } catch (err) {
    return json_({ok: false});
  }
}
