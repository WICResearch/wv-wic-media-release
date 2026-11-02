/**************************************************************
 * WV WIC MEDIA RELEASE REGISTRY
 * Backend - Code.gs
 **************************************************************/

const CONFIG = {
  DRIVE_FOLDER_ID: '1q7GzdycVJeN1yxaIVdkE8zFjLdGU6KLb',

  SHEETS: {
    RELEASES: 'Releases',
    EVENTS: 'Events',
    STAFF: 'Staff',
    AUDIT: 'Audit_Log'
  },

  RELEASE_PREFIX: 'WIC-MR',
  TIME_ZONE: 'America/New_York'
};


/* ============================================================
   WEB APP
============================================================ */

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('WV WIC Media Release Registry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}


/* ============================================================
   INITIAL SETUP
============================================================ */

function setupMediaReleaseRegistry() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const requiredSheets = [
    CONFIG.SHEETS.RELEASES,
    CONFIG.SHEETS.EVENTS,
    CONFIG.SHEETS.STAFF,
    CONFIG.SHEETS.AUDIT
  ];

  requiredSheets.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });

  setupEventsSheet_();
  setupStaffSheet_();
  setupAuditSheet_();

  // Verify Drive folder
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

  // Create subfolders if they don't exist
  getOrCreateFolder_(folder, 'Reference Photos');
  getOrCreateFolder_(folder, 'Signed Releases');

  logAudit_(
    'SYSTEM_SETUP',
    '',
    'Media Release Registry setup completed.'
  );

  return {
    success: true,
    message: 'WV WIC Media Release Registry setup complete.'
  };
}


/* ============================================================
   RELEASE CREATION
============================================================ */

function saveRelease(formData) {

  if (!formData) {
    throw new Error('No release information was received.');
  }

  validateRelease_(formData);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEETS.RELEASES);

    if (!sheet) {
      throw new Error('The Releases sheet could not be found.');
    }

    const headers = getHeaders_(sheet);

    const releaseId = generateReleaseId_(sheet);

    const now = new Date();
    const releaseDate = parseDate_(formData.releaseDate) || now;

    const retentionDate = calculateRetentionDate_(releaseDate);

    const filename = buildReleaseFilename_(
      formData.subjectRelationship,
      formData.subjectLastName,
      formData.subjectFirstName,
      releaseDate
    );


    /* --------------------------------------------------------
       SAVE REFERENCE PHOTO
    -------------------------------------------------------- */

    let referencePhotoUrl = '';

    if (formData.referencePhoto) {

      const referenceFilename =
        sanitizeFilename_(filename + '_Reference_Photo.jpg');

      referencePhotoUrl = saveBase64File_(
        formData.referencePhoto,
        'Reference Photos',
        referenceFilename
      );
    }


    /* --------------------------------------------------------
       SAVE SIGNED RELEASE IMAGE
    -------------------------------------------------------- */

    let signedReleaseUrl = '';

    if (formData.signedReleaseImage) {

      const releaseFilename =
        sanitizeFilename_(filename + '_Signed_Release.jpg');

      signedReleaseUrl = saveBase64File_(
        formData.signedReleaseImage,
        'Signed Releases',
        releaseFilename
      );
    }


    /* --------------------------------------------------------
       DETERMINE STATUS
    -------------------------------------------------------- */

    const permissionStatus =
      String(formData.permissionStatus || '')
        .trim()
        .toUpperCase();

    const paperFiled =
      formData.paperCopyFiled === true ||
      formData.paperCopyFiled === 'true';

    let recordStatus = determineRecordStatus_(
      permissionStatus,
      referencePhotoUrl,
      signedReleaseUrl,
      paperFiled
    );


    /* --------------------------------------------------------
       BUILD DATABASE RECORD
    -------------------------------------------------------- */

    const record = {

      Release_ID: releaseId,

      Subject_First_Name:
        clean_(formData.subjectFirstName),

      Subject_Last_Name:
        clean_(formData.subjectLastName),

      Subject_Relationship:
        clean_(formData.subjectRelationship),

      Is_Minor:
        formData.isMinor ? 'YES' : 'NO',

      Parent_Guardian_First_Name:
        clean_(formData.guardianFirstName),

      Parent_Guardian_Last_Name:
        clean_(formData.guardianLastName),

      Event_ID:
        clean_(formData.eventId),

      Event_Name:
        clean_(formData.eventName),

      Release_Date:
        releaseDate,

      Permission_Status:
        permissionStatus,

      Reference_Photo_URL:
        referencePhotoUrl,

      Signed_Release_URL:
        signedReleaseUrl,

      Paper_Copy_Filed:
        paperFiled ? 'YES' : 'NO',

      Paper_Copy_Filed_Date:
        paperFiled ? now : '',

      Staff_Name:
        clean_(formData.staffName),

      Retention_Until:
        retentionDate,

      Release_Filename:
        filename,

      Record_Status:
        recordStatus,

      Notes:
        clean_(formData.notes),

      Created_At:
        now,

      Updated_At:
        now
    };


    appendRecordByHeaders_(sheet, headers, record);


    /* --------------------------------------------------------
       AUDIT
    -------------------------------------------------------- */

    logAudit_(
      'RELEASE_CREATED',
      releaseId,
      'Release created. Status: ' + recordStatus
    );


    return {

      success: true,

      releaseId: releaseId,

      status: recordStatus,

      participant:
        clean_(formData.subjectFirstName) +
        ' ' +
        clean_(formData.subjectLastName),

      retentionUntil:
        Utilities.formatDate(
          retentionDate,
          CONFIG.TIME_ZONE,
          'MM/dd/yyyy'
        ),

      message:
        'Media release record saved successfully.'
    };

  } finally {

    lock.releaseLock();
  }
}


/* ============================================================
   STATUS LOGIC
============================================================ */

function determineRecordStatus_(
  permissionStatus,
  referencePhotoUrl,
  signedReleaseUrl,
  paperFiled
) {

  if (permissionStatus === 'DECLINED') {
    return 'DECLINED';
  }

  if (permissionStatus === 'REVOKED') {
    return 'REVOKED';
  }

  if (permissionStatus !== 'APPROVED') {
    return 'PENDING';
  }

  if (
    !referencePhotoUrl ||
    !signedReleaseUrl ||
    !paperFiled
  ) {
    return 'INCOMPLETE';
  }

  return 'APPROVED';
}


/* ============================================================
   APPROVED PARTICIPANT GALLERY
============================================================ */

function getApprovedParticipants(filters) {

  filters = filters || {};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RELEASES);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();

  const headers = values.shift();

  const records = values.map(row => {

    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = row[index];
    });

    return obj;
  });


  return records

    .filter(record => {

      if (
        String(record.Record_Status)
          .toUpperCase() !== 'APPROVED'
      ) {
        return false;
      }

      if (
        filters.eventName &&
        String(record.Event_Name)
          .toLowerCase() !==
        String(filters.eventName)
          .toLowerCase()
      ) {
        return false;
      }

      if (filters.search) {

        const search =
          String(filters.search)
            .toLowerCase();

        const fullName =
          (
            record.Subject_First_Name +
            ' ' +
            record.Subject_Last_Name
          ).toLowerCase();

        if (!fullName.includes(search)) {
          return false;
        }
      }

      return true;
    })

    .map(record => ({

      releaseId:
        record.Release_ID,

      firstName:
        record.Subject_First_Name,

      lastName:
        record.Subject_Last_Name,

      relationship:
        record.Subject_Relationship,

      eventName:
        record.Event_Name,

      releaseDate:
        formatDateForClient_(record.Release_Date),

      retentionUntil:
        formatDateForClient_(record.Retention_Until),

      referencePhotoUrl:
        record.Reference_Photo_URL,

      signedReleaseUrl:
        record.Signed_Release_URL,

      status:
        record.Record_Status

    }));
}


/* ============================================================
   SEARCH ALL RELEASES
============================================================ */

function searchReleases(searchText) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.RELEASES);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  const headers = data.shift();

  const query =
    String(searchText || '')
      .trim()
      .toLowerCase();


  return data

    .map(row => {

      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      return obj;
    })

    .filter(record => {

      if (!query) {
        return true;
      }

      const haystack = [
        record.Release_ID,
        record.Subject_First_Name,
        record.Subject_Last_Name,
        record.Parent_Guardian_First_Name,
        record.Parent_Guardian_Last_Name,
        record.Event_Name,
        record.Record_Status
      ]
        .join(' ')
        .toLowerCase();

      return haystack.includes(query);
    })

    .map(record => ({

      releaseId:
        record.Release_ID,

      participant:
        (
          record.Subject_First_Name +
          ' ' +
          record.Subject_Last_Name
        ).trim(),

      eventName:
        record.Event_Name,

      status:
        record.Record_Status,

      permissionStatus:
        record.Permission_Status,

      releaseDate:
        formatDateForClient_(record.Release_Date),

      referencePhotoUrl:
        record.Reference_Photo_URL,

      signedReleaseUrl:
        record.Signed_Release_URL

    }))

    .slice(0, 100);
}


/* ============================================================
   EVENTS
============================================================ */

function getEvents() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.EVENTS);

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const data = sheet.getDataRange().getValues();

  const headers = data.shift();

  return data

    .map(row => {

      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      return obj;
    })

    .filter(event =>
      String(event.Active)
        .toUpperCase() !== 'NO'
    )

    .map(event => ({

      id: event.Event_ID,

      name: event.Event_Name,

      date:
        formatDateForClient_(event.Event_Date),

      location:
        event.Location

    }));
}


function addEvent(eventData) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.EVENTS);

  const now = new Date();

  const eventId =
    'EVT-' +
    Utilities.formatDate(
      now,
      CONFIG.TIME_ZONE,
      'yyyyMMdd-HHmmss'
    );

  sheet.appendRow([
    eventId,
    clean_(eventData.name),
    parseDate_(eventData.date),
    clean_(eventData.location),
    'YES',
    now
  ]);

  logAudit_(
    'EVENT_CREATED',
    eventId,
    clean_(eventData.name)
  );

  return {
    success: true,
    eventId: eventId
  };
}


/* ============================================================
   FILE UPLOADS
============================================================ */

function saveBase64File_(
  dataUrl,
  subfolderName,
  filename
) {

  if (!dataUrl) {
    return '';
  }

  const match =
    String(dataUrl)
      .match(/^data:(.+);base64,(.+)$/);

  if (!match) {
    throw new Error(
      'The uploaded image was not in a valid format.'
    );
  }

  const mimeType = match[1];

  const bytes =
    Utilities.base64Decode(match[2]);

  const blob =
    Utilities.newBlob(
      bytes,
      mimeType,
      filename
    );

  const root =
    DriveApp.getFolderById(
      CONFIG.DRIVE_FOLDER_ID
    );

  const folder =
    getOrCreateFolder_(
      root,
      subfolderName
    );

  const file =
    folder.createFile(blob);

  return file.getUrl();
}


/* ============================================================
   RELEASE ID
============================================================ */

function generateReleaseId_(sheet) {

  const year =
    Utilities.formatDate(
      new Date(),
      CONFIG.TIME_ZONE,
      'yyyy'
    );

  const prefix =
    CONFIG.RELEASE_PREFIX +
    '-' +
    year +
    '-';

  const lastRow =
    sheet.getLastRow();

  let highest = 0;

  if (lastRow >= 2) {

    const ids =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          1
        )
        .getValues()
        .flat();

    ids.forEach(id => {

      const value =
        String(id || '');

      if (value.startsWith(prefix)) {

        const number =
          parseInt(
            value.replace(prefix, ''),
            10
          );

        if (
          !isNaN(number) &&
          number > highest
        ) {
          highest = number;
        }
      }
    });
  }

  const next =
    String(highest + 1)
      .padStart(5, '0');

  return prefix + next;
}


/* ============================================================
   RETENTION
   Release date + 3 years + 150 days
============================================================ */

function calculateRetentionDate_(releaseDate) {

  const date =
    new Date(releaseDate);

  date.setFullYear(
    date.getFullYear() + 3
  );

  date.setDate(
    date.getDate() + 150
  );

  return date;
}


/* ============================================================
   FILE NAMING
============================================================ */

function buildReleaseFilename_(
  relationship,
  lastName,
  firstName,
  releaseDate
) {

  const dateString =
    Utilities.formatDate(
      releaseDate,
      CONFIG.TIME_ZONE,
      'MM.dd.yyyy'
    );

  return sanitizeFilename_(
    clean_(relationship) +
    '_' +
    clean_(lastName) +
    '_' +
    clean_(firstName) +
    '_' +
    dateString
  );
}


function sanitizeFilename_(filename) {

  return String(filename || '')
    .replace(/[\\\/:*?"<>|#%{}~]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


/* ============================================================
   VALIDATION
============================================================ */

function validateRelease_(data) {

  if (!clean_(data.subjectFirstName)) {
    throw new Error(
      'Subject first name is required.'
    );
  }

  if (!clean_(data.subjectLastName)) {
    throw new Error(
      'Subject last name is required.'
    );
  }

  if (!clean_(data.subjectRelationship)) {
    throw new Error(
      'Subject relationship is required.'
    );
  }

  if (!clean_(data.releaseDate)) {
    throw new Error(
      'Release date is required.'
    );
  }

  if (!clean_(data.permissionStatus)) {
    throw new Error(
      'Permission status is required.'
    );
  }

  if (data.isMinor) {

    if (
      !clean_(data.guardianFirstName) ||
      !clean_(data.guardianLastName)
    ) {
      throw new Error(
        'Parent/guardian information is required for a minor.'
      );
    }
  }
}


/* ============================================================
   SHEET HELPERS
============================================================ */

function getHeaders_(sheet) {

  return sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .getValues()[0]
    .map(header =>
      String(header).trim()
    );
}


function appendRecordByHeaders_(
  sheet,
  headers,
  record
) {

  const row =
    headers.map(header => {

      if (
        Object.prototype.hasOwnProperty.call(
          record,
          header
        )
      ) {
        return record[header];
      }

      return '';
    });

  sheet.appendRow(row);
}


/* ============================================================
   SUPPORTING SHEETS
============================================================ */

function setupEventsSheet_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.EVENTS
      );

  if (sheet.getLastRow() === 0) {

    sheet.appendRow([
      'Event_ID',
      'Event_Name',
      'Event_Date',
      'Location',
      'Active',
      'Created_At'
    ]);

    styleHeader_(sheet);
  }
}


function setupStaffSheet_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.STAFF
      );

  if (sheet.getLastRow() === 0) {

    sheet.appendRow([
      'Staff_Email',
      'Staff_Name',
      'Role',
      'Active',
      'Created_At'
    ]);

    styleHeader_(sheet);
  }
}


function setupAuditSheet_() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.AUDIT
      );

  if (sheet.getLastRow() === 0) {

    sheet.appendRow([
      'Timestamp',
      'User_Email',
      'Action',
      'Record_ID',
      'Details'
    ]);

    styleHeader_(sheet);
  }
}


/* ============================================================
   AUDIT LOG
============================================================ */

function logAudit_(
  action,
  recordId,
  details
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        CONFIG.SHEETS.AUDIT
      );

  if (!sheet) {
    return;
  }

  let email = '';

  try {
    email =
      Session
        .getActiveUser()
        .getEmail();
  } catch (error) {
    email = '';
  }

  sheet.appendRow([
    new Date(),
    email,
    action,
    recordId || '',
    details || ''
  ]);
}


/* ============================================================
   DRIVE HELPERS
============================================================ */

function getOrCreateFolder_(
  parent,
  name
) {

  const folders =
    parent.getFoldersByName(name);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parent.createFolder(name);
}


/* ============================================================
   FORMATTING
============================================================ */

function styleHeader_(sheet) {

  if (sheet.getLastColumn() === 0) {
    return;
  }

  sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .setFontWeight('bold');

  sheet.setFrozenRows(1);
}


function formatDateForClient_(value) {

  if (!value) {
    return '';
  }

  const date =
    value instanceof Date
      ? value
      : new Date(value);

  if (isNaN(date.getTime())) {
    return String(value);
  }

  return Utilities.formatDate(
    date,
    CONFIG.TIME_ZONE,
    'MM/dd/yyyy'
  );
}


function parseDate_(value) {

  if (!value) {
    return null;
  }

  if (value instanceof Date) {
    return value;
  }

  const parts =
    String(value)
      .split('-');

  if (parts.length === 3) {

    return new Date(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2])
    );
  }

  const parsed =
    new Date(value);

  return isNaN(parsed.getTime())
    ? null
    : parsed;
}


function clean_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }

  return String(value).trim();
}
