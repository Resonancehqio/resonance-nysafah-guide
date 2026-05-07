// ── NYSAFAH Guide → Google Sheets Sync ──
// Paste this entire file into Extensions > Apps Script in your Google Sheet.
// Deploy as web app: Execute as Me, Anyone can access.
// After updating this code, go to Deploy > Manage deployments > Edit > New version > Deploy.

const SHEET_TOKEN = 'nysafah2026resonance';

function doGet(e) {
  try {
    const token = e.parameter.token || '';
    if (token !== SHEET_TOKEN) return jsonResponse({ error: 'Unauthorized' });

    // Write mode — single contact update sent as URL params
    if (e.parameter.write === '1') {
      const id = e.parameter.id;
      if (!id) return jsonResponse({ error: 'Missing id' });

      const sheet = getOrCreateSheet();
      const existing = sheet.getDataRange().getValues();
      const rowIndex = {};
      for (let i = 1; i < existing.length; i++) {
        if (existing[i][0]) rowIndex[existing[i][0]] = i + 1;
      }

      const row = [
        id,
        e.parameter.name   || '',
        e.parameter.org    || '',
        e.parameter.tier   || '',
        e.parameter.spoke === '1' ? 'Yes' : 'No',
        e.parameter.time   || '',
        e.parameter.day    || '',
        e.parameter.action || '',
        e.parameter.notes  || '',
        new Date().toISOString()
      ];

      if (rowIndex[id]) {
        sheet.getRange(rowIndex[id], 1, 1, row.length).setValues([row]);
      } else {
        sheet.appendRow(row);
      }

      return jsonResponse({ success: true });
    }

    // Read mode — return all contacts as JSON
    const sheet = getOrCreateSheet();
    if (sheet.getLastRow() < 2) return jsonResponse({ contacts: {} });

    const data = sheet.getDataRange().getValues();
    const contacts = {};
    for (let i = 1; i < data.length; i++) {
      const [id, name, org, tier, spoke, time, day, action, notes] = data[i];
      if (!id) continue;
      contacts[id] = { name, org, tier, spoke: spoke === 'Yes', time: time || '', day: day || 1, action: action || '', notes: notes || '' };
    }
    return jsonResponse({ contacts });

  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Contacts');
  if (!sheet) {
    sheet = ss.insertSheet('Contacts');
    sheet.appendRow(['ID', 'Name', 'Org', 'Tier', 'Spoke To', 'Time (ISO)', 'Day', 'Next Action', 'Notes', 'Last Synced']);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(9, 300);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
