// Pravin & Saranya wedding — guest photo album backend (Google Apps Script)
// Deploy: script.google.com > New project > paste this > Deploy > New deployment
//         Type: Web app · Execute as: Me · Who has access: Anyone
var FOLDER_ID = '1Gac2OpZJZBdPxEUolfGXSy7r6qHLeg96';
var MAX_BYTES = 6 * 1024 * 1024;
var MAX_LIST  = 60;

function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}

// List the newest photos (ids only; the page builds thumbnail URLs from them)
function doGet(){
  var it = DriveApp.getFolderById(FOLDER_ID).getFiles(), out = [];
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType().indexOf('image/') === 0) {
      if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);  // photos added by hand
      out.push({id: f.getId(), t: f.getDateCreated().getTime()});
    }
  }
  out.sort(function(a,b){return b.t - a.t;});
  return json_(out.slice(0, MAX_LIST));
}

// Receive one photo as {type, data(base64)} and store it in the folder
function doPost(e){
  try {
    var d = JSON.parse(e.postData.contents);
    if (!/^image\/(jpeg|png|webp)$/.test(d.type)) throw new Error('type');
    var bytes = Utilities.base64Decode(d.data);
    if (bytes.length > MAX_BYTES) throw new Error('size');
    var ext = d.type === 'image/png' ? 'png' : d.type === 'image/webp' ? 'webp' : 'jpg';
    var file = DriveApp.getFolderById(FOLDER_ID).createFile(Utilities.newBlob(bytes, d.type, 'guest-' + Date.now() + '.' + ext));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   // so thumbnails can show on the site
    return json_({ok: true, id: file.getId()});
  } catch (err) {
    return json_({ok: false});
  }
}
