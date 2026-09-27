/**
 * SISTEM PENJANA LAPORAN ONE PAGE REPORT (OPR) - SMK TAMAN NUSA JAYA
 * Google Apps Script Backend TANPA AI
 * Fungsi: read, create, update, delete dan upload gambar ke Google Drive.
 */

const DEFAULT_SHEET_ID = '1rVUp83WjhNLfuuxh-CW3PhEVQXTlD5ch0PKvSgptOuE';
const DEFAULT_FOLDER_ID = '1iYKu-k_i32uvPr2dmOkCjV3U020pCwnW';

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || '';
    const sheetId = params.sheetId || DEFAULT_SHEET_ID;

    if (action === 'read') {
      const ss = SpreadsheetApp.openById(sheetId);
      const sheet = ss.getActiveSheet();
      const data = sheet.getDataRange().getValues();

      if (data.length <= 1) {
        return createJsonResponse({ status: 'success', data: [] });
      }

      const reports = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0] && !row[2]) continue;

        reports.push({
          id: row[0] || '',
          timestamp: formatSheetValue(row[1]),
          namaProgram: row[2] || '',
          unit: row[3] || '',
          tarikhMula: formatDateForInput(row[4]),
          tarikhAkhir: formatDateForInput(row[5]),
          tempat: row[6] || '',
          sasaran: row[7] || '',
          objektif: row[8] || '',
          kekuatan: row[9] || '',
          kelemahan: row[10] || '',
          penambahbaikan: row[11] || '',
          namaGuru: row[12] || '',
          jawatanGuru: row[13] || '',
          images: [row[14] || '', row[15] || '', row[16] || '']
        });
      }

      return createJsonResponse({ status: 'success', data: reports });
    }

    return createJsonResponse({ status: 'error', message: 'Parameter tidak sah' });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Tiada data POST diterima.');
    }

    const postData = JSON.parse(e.postData.contents);
    const action = postData.action || '';
    const sheetId = postData.sheetId || DEFAULT_SHEET_ID;
    const folderId = postData.folderId || DEFAULT_FOLDER_ID;
    const report = postData.report;

    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getActiveSheet();
    ensureHeader(sheet);

    if (action === 'create' || action === 'update') {
      if (!report) throw new Error('Data laporan tidak diberikan.');

      const imageUrls = ['', '', ''];
      if (report.images && Array.isArray(report.images)) {
        for (let idx = 0; idx < Math.min(report.images.length, 3); idx++) {
          const imgStr = report.images[idx];
          if (imgStr && imgStr.indexOf('data:image') === 0) {
            imageUrls[idx] = uploadBase64ToDrive(
              imgStr,
              folderId,
              report.id + '_img' + (idx + 1)
            );
          } else if (imgStr) {
            imageUrls[idx] = imgStr;
          }
        }
      }

      const rowValues = [
        report.id,
        report.timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        report.namaProgram || '',
        report.unit || '',
        report.tarikhMula || '',
        report.tarikhAkhir || '',
        report.tempat || '',
        report.sasaran || '',
        report.objektif || '',
        report.kekuatan || '',
        report.kelemahan || '',
        report.penambahbaikan || '',
        report.namaGuru || '',
        report.jawatanGuru || '',
        imageUrls[0],
        imageUrls[1],
        imageUrls[2]
      ];

      if (action === 'update') {
        const data = sheet.getDataRange().getValues();
        let foundRow = -1;

        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0]) === String(report.id)) {
            foundRow = i + 1;
            break;
          }
        }

        if (foundRow !== -1) {
          sheet.getRange(foundRow, 1, 1, rowValues.length).setValues([rowValues]);
        } else {
          sheet.appendRow(rowValues);
        }
      } else {
        sheet.appendRow(rowValues);
      }

      return createJsonResponse({
        status: 'success',
        message: action === 'update' ? 'Laporan berjaya dikemas kini!' : 'Laporan berjaya disimpan!',
        id: report.id
      });
    }

    if (action === 'delete') {
      const idToDelete = postData.id;
      if (!idToDelete) throw new Error('ID laporan tidak diberikan.');

      const data = sheet.getDataRange().getValues();
      let deleted = false;

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idToDelete)) {
          sheet.deleteRow(i + 1);
          deleted = true;
          break;
        }
      }

      return createJsonResponse({
        status: 'success',
        message: deleted ? 'Laporan berjaya dipadam!' : 'Rekod tidak dijumpai.'
      });
    }

    return createJsonResponse({ status: 'error', message: 'Aksi tidak sah' });
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.message || String(err) });
  }
}

function ensureHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'ID Laporan', 'Tarikh Cipta', 'Nama Program', 'Unit', 'Tarikh Mula',
      'Tarikh Akhir', 'Tempat', 'Sasaran Peserta', 'Objektif', 'Kekuatan',
      'Kelemahan', 'Penambahbaikan', 'Nama Guru', 'Jawatan Guru',
      'Gambar 1 (URL)', 'Gambar 2 (URL)', 'Gambar 3 (URL)'
    ]);

    sheet.getRange(1, 1, 1, 17)
      .setFontWeight('bold')
      .setBackground('#e2e8f0');
    sheet.setFrozenRows(1);
  }
}

function uploadBase64ToDrive(base64Data, folderId, filename) {
  const splitData = base64Data.split(',');
  if (splitData.length < 2) throw new Error('Format gambar tidak sah.');

  const mimeMatch = splitData[0].match(/:(.*?);/);
  if (!mimeMatch || !mimeMatch[1]) throw new Error('Jenis gambar tidak dapat dikenal pasti.');

  const contentType = mimeMatch[1];
  const extensionMap = {
    'image/jpeg': '.jpg',
    'image/jpg': '.jpg',
    'image/png': '.png',
    'image/webp': '.webp',
    'image/gif': '.gif'
  };

  const bytes = Utilities.base64Decode(splitData[1]);
  const blob = Utilities.newBlob(
    bytes,
    contentType,
    filename + (extensionMap[contentType] || '')
  );

  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);

  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (sharingError) {
    Logger.log('Sharing warning: ' + sharingError);
  }

  return file.getUrl();
}

function formatSheetValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return value == null ? '' : String(value);
}

function formatDateForInput(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  if (value == null) return '';
  const str = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) return str;
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(str)) {
    const p = str.split('/');
    return `${p[2]}-${p[1]}-${p[0]}`;
  }
  return str;
}

function createJsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
