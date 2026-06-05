var SHEET_ID = '1RSlchKcmZ1bgcV2S8vM8dTl-1maG4BHExArq8mCLAUM';

function doGet(e) {
  var type = e.parameter.type || '';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var result;

  if (type === 'active') {
    var sheet = ss.getSheetByName('Active Match');
    if (!sheet) {
      result = { active: false };
    } else {
      var data = sheet.getRange(2, 1, 1, 9).getValues()[0];
      if (!data[0]) {
        result = { active: false };
      } else {
        result = {
          active:         true,
          league:         data[0],
          homeTeam:       data[1],
          awayTeam:       data[2],
          gameDate:       data[3],
          jackpot:        data[4],
          matchId:        data[5],
          winTeam:        data[6],
          matchStartTime: data[7],
          cutoffTime:     data[8]
        };
      }
    }
  }

  else if (type === 'entries') {
    var sheet = ss.getSheetByName('Entries');
    if (!sheet) {
      result = [];
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        result = [];
      } else {
        var rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
        var filtered = rows.filter(function(r) { return r[0]; });
        result = filtered.map(function(r) {
          return {
            timestamp: r[0],
            name:      r[1],
            contact:   r[2],
            matchId:   r[3],
            score:     r[4],
            league:    r[5],
            email:     r[6]
          };
        });
      }
    }
  }

  else if (type === 'history') {
    var sheet = ss.getSheetByName('Match History');
    if (!sheet) {
      result = [];
    } else {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) {
        result = [];
      } else {
        var rows = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
        result = rows.filter(function(r) { return r[0]; }).map(function(r) {
          return {
            date:     r[0],
            league:   r[1],
            home:     r[2],
            away:     r[3],
            winTeam:  r[4],
            winScore: r[5],
            winners:  r[6],
            closest:  r[7],
            prize:    r[8]
          };
        });
      }
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var data = JSON.parse(e.postData.contents);
  var type = data.type || '';

  if (type === 'setMatch' || type === 'setActive') {
    var sheet = ss.getSheetByName('Active Match');
    if (!sheet) sheet = ss.insertSheet('Active Match');
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 9).setValues([['League','Home Team','Away Team','Game Date','Jackpot','Match ID','Win Team','Match Start Time','Cutoff Time']]);
    var matchId = data.matchId || Date.now();
    sheet.getRange(2, 1, 1, 9).setValues([[
      data.league      || '',
      data.home        || '',
      data.away        || '',
      data.date        || '',
      data.jackpot     || data.prize || 1000000,
      matchId,
      '',
      matchId,
      data.cutoffTime  || ''
    ]]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, matchId: matchId }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (type === 'entry') {
    var sheet = ss.getSheetByName('Entries');
    if (!sheet) sheet = ss.insertSheet('Entries');
    if (sheet.getLastRow() < 1) {
      sheet.getRange(1, 1, 1, 7).setValues([['Timestamp','Name','Contact','Match ID','Score','League','Email']]);
    }
    sheet.appendRow([
      new Date(),
      data.name    || '',
      data.contact || '',
      data.matchId || '',
      data.score   || '',
      data.league  || '',
      data.email   || ''
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (type === 'match' || type === 'result') {
    var histSheet = ss.getSheetByName('Match History');
    if (!histSheet) histSheet = ss.insertSheet('Match History');
    if (histSheet.getLastRow() < 1) {
      histSheet.getRange(1, 1, 1, 9).setValues([['Date','League','Home','Away','Win Team','Win Score','Winners','Closest Entry','Prize Pool']]);
    }
    histSheet.appendRow([
      new Date(),
      data.league   || '',
      data.home     || '',
      data.away     || '',
      data.winTeam  || '',
      data.winScore || '',
      data.winners  || 'None',
      data.closest  || '',
      data.prize    || ''
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (type === 'clearActive') {
    var activeSheet = ss.getSheetByName('Active Match');
    if (activeSheet) {
      activeSheet.getRange(2, 1, 1, 9).clearContent();
    }
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: 'Unknown type' }))
    .setMimeType(ContentService.MimeType.JSON);
}
