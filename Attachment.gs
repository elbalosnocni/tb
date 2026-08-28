/**
 * Attachment storage in Google Drive.
 * Creates a folder "CompanyNoticeAttachments" in My Drive.
 * The web app stores only metadata/URL in Announcements.
 */
function uploadAttachment(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  if (!payload || !payload.base64 || !payload.name) throw new Error('Thiếu tệp.');

  const bytes = Utilities.base64Decode(payload.base64);
  if (bytes.length > CONFIG.maxAttachmentBytes) throw new Error('Tệp vượt quá 25MB.');

  const folder = getAttachmentFolder_();
  const safeName = String(payload.name).replace(/[\\/:*?"<>|#%]/g,'_');
  const blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', safeName);
  const file = folder.createFile(blob);

  // "Anyone with the link" may be restricted by Workspace admin policy.
  // If forbidden, the preview will require an authorized Google account.
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}

  const id = file.getId();
  const previewUrl = 'https://drive.google.com/file/d/' + id + '/preview';
  const url = 'https://drive.google.com/uc?export=download&id=' + id;

  audit_(session.employeeId,'UPLOAD_ATTACHMENT','DriveFile',id,safeName);

  return {
    ok:true,
    attachment:{
      id:id,name:safeName,mimeType:file.getMimeType(),size:bytes.length,
      url:url,previewUrl:previewUrl
    }
  };
}

function getAttachmentFolder_() {
  const props=PropertiesService.getScriptProperties();
  const existing=props.getProperty('ATTACHMENT_FOLDER_ID');
  if(existing) {
    try{return DriveApp.getFolderById(existing)}catch(e){}
  }
  const it=DriveApp.getFoldersByName('CompanyNoticeAttachments');
  const folder=it.hasNext()?it.next():DriveApp.createFolder('CompanyNoticeAttachments');
  props.setProperty('ATTACHMENT_FOLDER_ID',folder.getId());
  return folder;
}
