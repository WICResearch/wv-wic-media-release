/**
 * DigitalRelease.gs — ADD AS A NEW FILE to your existing Apps Script project.
 * Prototype only. Do not use with real participant data until WIC policy,
 * records management, privacy/security, and IT approvals are complete.
 * This does not change your existing Releases sheet or existing saveRelease().
 */
const DIGITAL_RELEASE_PROTO = {
  FOLDER_ID: '1q7GzdycVJeN1yxaIVdkE8zFjLdGU6KLb',
  TAB: 'Digital_Prototype',
  VERSION: 'Policy 7.02 Attachment #1 Revision 12/03/25',
  CONSENT: ["I hereby grant the West Virginia WIC Program and their licenses, agents, successors and assigns the right to use my appearance and likeness in photograph(s), videotape(s) and voice recording(s) in their products, catalogs, books, educational products, audiovisual materials, promotional material and publications.", "I hereby grant the West Virginia WIC Program and their licenses, agents, successors and assigns the right to use my written comments and name in promotional material and publications.", "On my own behalf, and on the behalf of my heirs, next of kin, executors, administrators, successors, and assigns, I hereby release the West Virginia WIC Program, its employees, the director and all their agents from any and all claim, liability and damages arising out of the rights granted hereunder, or the exercise thereof."],
  USDA: "In accordance with Federal civil rights law and U.S. Department of Agriculture (USDA) civil rights regulations and policies, the USDA, its Agencies, offices, and employees, and institutions participating in or administering USDA programs are prohibited from discriminating based on race, color, national origin, religion, sex, disability, age, marital status, family/parental status, income derived from a public assistance program, political beliefs, or reprisal or retaliation for prior civil rights activity, in any program or activity conducted or funded by USDA (not all bases apply to all programs). Remedies and complaint filing deadlines vary by program or incident. Persons with disabilities who require alternative means of communication for program information (e.g., Braille, large print, audiotape, American Sign Language, etc.) should contact the State or local Agency that administers the program or contact USDA through the Telecommunications Relay Service at 711 (voice and TTY). Additionally, program information may be made available in languages other than English. To file a program discrimination complaint, complete the USDA Program Discrimination Complaint Form, AD-3027, found online at How to File a Program Discrimination Complaint and Policy 7.02 Attachment #1 Revision: 12/03/25 at any USDA office or write a letter addressed to USDA and provide in the letter all of the information requested in the form. To request a copy of the complaint form, call (866) 632-9992. Submit your completed form or letter to USDA by: (1) mail: U.S. Department of Agriculture, Office of the Assistant Secretary for Civil Rights, 1400 Independence Avenue, SW, Mail Stop 9410, Washington, D.C. 20250-9410; (2) fax: (202) 690-7442; or (3) email: program.intake@usda.gov. USDA is an equal opportunity provider, employer, and lender."
};

function saveDigitalRelease(data) {
  if (!data || !data.subjectFirstName || !data.subjectLastName || !data.releaseDate ||
      !data.staffName || !data.signerAddress || !data.signerPhone ||
      !['YES','NO'].includes(data.signerOver18) ||
      !['APPROVED','DECLINED'].includes(data.permissionStatus) ||
      data.consentAcknowledged !== true || !data.signatureImage) {
    throw new Error('Missing required digital release information.');
  }
  if (!['STATE','LOCAL'].includes(data.agencyType) || !String(data.agencyName || '').trim() ||
      (data.agencyType === 'STATE' && data.agencyName !== 'State WIC Office')) {
    throw new Error('A valid staff agency selection is required.');
  }
  if (data.isMinor && (!data.guardianFirstName || !data.guardianLastName || data.signerOver18 !== 'YES')) {
    throw new Error('An adult parent/guardian is required for a minor.');
  }
  if (data.permissionStatus === 'APPROVED' && !data.referencePhoto) {
    throw new Error('A reference photo is required for this prototype.');
  }
  if (data.consentVersion !== DIGITAL_RELEASE_PROTO.VERSION) {
    throw new Error('The consent version does not match the current prototype.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('No bound spreadsheet found. Open this script from the registry spreadsheet.');
    let sheet = ss.getSheetByName(DIGITAL_RELEASE_PROTO.TAB);
    if (!sheet) sheet = ss.insertSheet(DIGITAL_RELEASE_PROTO.TAB);
    const headers = ['Record_ID','Subject_First_Name','Subject_Last_Name','Subject_Relationship',
      'Is_Minor','Guardian_First_Name','Guardian_Last_Name','Event_ID','Event_Name','Release_Date',
      'Consent_Decision','Signer_Address','Signer_Phone','Signer_Over_18','Consent_Version',
      'Consent_Acknowledged','Signature_Captured','Staff_Name','Notes','PDF_URL',
      'Reference_Photo_URL','Submitted_At','Record_Status'];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);
    const existingHeaders = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    ['Agency_Type','Agency_Name'].forEach(function(h){
      if (existingHeaders.indexOf(h) === -1) {
        sheet.getRange(1,sheet.getLastColumn()+1).setValue(h);
        existingHeaders.push(h);
      }
    });
    const folder = DriveApp.getFolderById(DIGITAL_RELEASE_PROTO.FOLDER_ID);
    const subfolder = digitalProtoFolder_(folder,'Digital Prototype — TEST ONLY');
    const now = new Date();
    const id = 'PROTOTYPE-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0,6);
    const subject = String(data.subjectFirstName).trim() + ' ' + String(data.subjectLastName).trim();
    const doc = DocumentApp.create(id + ' — ' + subject);
    const body = doc.getBody();
    body.appendParagraph('WV WIC MEDIA RELEASE — DIGITAL PROTOTYPE').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('NOT AN AUTHORIZED SUBSTITUTE FOR POLICY 7.02 PAPER RELEASE');
    body.appendParagraph('Record ID: ' + id);
    body.appendParagraph('Consent document: ' + DIGITAL_RELEASE_PROTO.VERSION);
    body.appendParagraph('Photo Subject(s) Name(s): ' + subject);
    body.appendParagraph('Subject Parent/Guardian Name (if applicable): ' + [data.guardianFirstName,data.guardianLastName].filter(Boolean).join(' '));
    body.appendParagraph('I am over 18 years of age: ' + data.signerOver18);
    body.appendParagraph('Address: ' + String(data.signerAddress));
    body.appendParagraph('Telephone Number: ' + String(data.signerPhone));
    body.appendParagraph('Release date: ' + String(data.releaseDate));
    body.appendParagraph('Event: ' + String(data.eventName || 'General release'));
    body.appendParagraph('Recorded authorization decision: ' + String(data.permissionStatus));
    body.appendParagraph('RELEASE FORM').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    DIGITAL_RELEASE_PROTO.CONSENT.forEach(function(p){body.appendParagraph(p);});
    body.appendParagraph('Subject or Parent/Guardian Signature (electronic prototype):');
    const signature = digitalProtoBlob_(data.signatureImage,'signature.png','image/png');
    body.appendImage(signature).setWidth(250);
    body.appendParagraph('Consent acknowledgement: ' + (data.consentAcknowledged ? 'Confirmed' : 'Not confirmed'));
    body.appendParagraph('Signed/submitted at: ' + now.toISOString());
    body.appendParagraph('Staff member: ' + String(data.staffName));
    body.appendParagraph('Submitting agency: ' + String(data.agencyName));
    body.appendParagraph('USDA NONDISCRIMINATION STATEMENT').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(DIGITAL_RELEASE_PROTO.USDA);
    doc.saveAndClose();
    const sourceFile = DriveApp.getFileById(doc.getId());
    const pdf = subfolder.createFile(sourceFile.getAs(MimeType.PDF).setName(id + '.pdf'));
    sourceFile.setTrashed(true);
    let photoUrl = '';
    if (data.permissionStatus === 'APPROVED' && data.referencePhoto) {
      const photo = subfolder.createFile(digitalProtoBlob_(data.referencePhoto,id + '-reference.jpg','image/jpeg'));
      photoUrl = photo.getUrl();
    }
    const recordValues = [id,data.subjectFirstName,data.subjectLastName,data.subjectRelationship,
      !!data.isMinor,data.guardianFirstName,data.guardianLastName,data.eventId,data.eventName,
      data.releaseDate,data.permissionStatus,data.signerAddress,data.signerPhone,data.signerOver18,
      DIGITAL_RELEASE_PROTO.VERSION,true,true,data.staffName,data.notes,pdf.getUrl(),photoUrl,
      now,'PROTOTYPE_NOT_VALID_FOR_MEDIA_USE'];
    const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const row = Array(headerRow.length).fill('');
    headers.forEach(function(h,i){const col=headerRow.indexOf(h);if(col>=0)row[col]=recordValues[i];});
    row[headerRow.indexOf('Agency_Type')] = data.agencyType;
    row[headerRow.indexOf('Agency_Name')] = String(data.agencyName).trim();
    sheet.appendRow(row);
    return {success:true,releaseId:id,participant:subject,status:'PROTOTYPE',
      retentionUntil:'Pending approved records schedule',pdfUrl:pdf.getUrl()};
  } finally { lock.releaseLock(); }
}

function digitalProtoFolder_(parent,name) {
  const iter=parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parent.createFolder(name);
}
function digitalProtoBlob_(dataUrl,name,expectedMime) {
  const match=String(dataUrl||'').match(/^data:(image\/(?:png|jpeg));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('Invalid image data for ' + name);
  const bytes=Utilities.base64Decode(match[2]);
  if (bytes.length > 8*1024*1024) throw new Error('Image is too large.');
  return Utilities.newBlob(bytes,match[1],name);
}
