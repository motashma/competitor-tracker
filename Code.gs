/**
 * Competitor Price & Content Tracker — Sheet Backend
 *
 * SETUP:
 * 1. Open your Google Sheet (tabs named exactly "Competitors" and "ChangeLog").
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code, paste this whole file in, save.
 * 4. Click Deploy > New deployment > type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL it gives you — that's the single URL both
 *    n8n and the dashboard will call.
 *
 * SHEET COLUMNS (row 1 = headers, exact names):
 * Competitors: ID | Name | URL | TrackType | Selector | LastChecked | LastHash | LastSnapshot
 * ChangeLog:   ID | CompetitorID | CompetitorName | Timestamp | Summary
 */

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function niceTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy 'at' h:mm a");
}

function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  return rows
    .filter(function (r) { return r.join('') !== ''; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = (e.parameter.action || 'list');
  if (action === 'list') {
    var competitors = sheetToObjects(getSheet('Competitors'));
    var changelog = sheetToObjects(getSheet('ChangeLog')).sort(function (a, b) {
      return new Date(b.Timestamp) - new Date(a.Timestamp);
    });
    return respond({ success: true, competitors: competitors, changelog: changelog });
  }
  return respond({ success: false, error: 'Unknown action' });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return respond({ success: false, error: 'Invalid JSON body' });
  }

  var action = data.action;
  var sheet = getSheet('Competitors');
  var logSheet = getSheet('ChangeLog');

  if (action === 'addCompetitor') {
    var id = 'C' + new Date().getTime();
    sheet.appendRow([
      id, data.name, data.url, data.trackType || 'full',
      data.selector || '', '', '', ''
    ]);
    return respond({ success: true, id: id });
  }

  if (action === 'updateCompetitor') {
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        if (data.name !== undefined) sheet.getRange(i + 1, 2).setValue(data.name);
        if (data.url !== undefined) sheet.getRange(i + 1, 3).setValue(data.url);
        if (data.trackType !== undefined) sheet.getRange(i + 1, 4).setValue(data.trackType);
        if (data.selector !== undefined) sheet.getRange(i + 1, 5).setValue(data.selector);
        return respond({ success: true });
      }
    }
    return respond({ success: false, error: 'Competitor not found' });
  }

  if (action === 'deleteCompetitor') {
    var rows2 = sheet.getDataRange().getValues();
    for (var j = 1; j < rows2.length; j++) {
      if (rows2[j][0] === data.id) {
        sheet.deleteRow(j + 1);
        return respond({ success: true });
      }
    }
    return respond({ success: false, error: 'Competitor not found' });
  }

  // Called by n8n after checking a page, whether or not it changed
  if (action === 'touch') {
    var rows3 = sheet.getDataRange().getValues();
    for (var k = 1; k < rows3.length; k++) {
      if (rows3[k][0] === data.id) {
        sheet.getRange(k + 1, 6).setValue(niceTimestamp());
        if (data.hash) sheet.getRange(k + 1, 7).setValue(data.hash);
        if (data.snapshot) sheet.getRange(k + 1, 8).setValue(String(data.snapshot).slice(0, 5000));
        return respond({ success: true });
      }
    }
    return respond({ success: false, error: 'Competitor not found' });
  }

  // Called by n8n when a real change was detected + summarized
  if (action === 'recordChange') {
    var rows4 = sheet.getDataRange().getValues();
    for (var m = 1; m < rows4.length; m++) {
      if (rows4[m][0] === data.competitorId) {
        sheet.getRange(m + 1, 6).setValue(niceTimestamp());
        sheet.getRange(m + 1, 7).setValue(data.hash);
        sheet.getRange(m + 1, 8).setValue(String(data.snapshot || '').slice(0, 5000));
        break;
      }
    }
    logSheet.appendRow([
      'L' + new Date().getTime(),
      data.competitorId,
      data.competitorName || '',
      niceTimestamp(),
      data.summary || ''
    ]);
    return respond({ success: true });
  }

  return respond({ success: false, error: 'Unknown action: ' + action });
}
