// معرّف Google Sheet الجديد واسم صفحة البيانات
const SHEET_ID = '1Xcofz05IsgKCqpmoh3-Bywm0WjX2FDFhCmWt_x7Vdr4';
const SHEET_NAME = 'بيانات';

// ترتيب الأعمدة:
// 0=name, 1=grand, 2=avg, 3=filled,
// 4-13=dayTotals(10), 14=dayBase(JSON), 15=dayBonus(JSON),
// 16=state, 17=lastUpdate, 18=notes, 19=editedPts,
// 20=counter, 21=personalGoals
const HEADERS = [
  'الاسم', 'المجموع', 'المعدل', 'الأيام',
  'اليوم 1', 'اليوم 2', 'اليوم 3', 'اليوم 4', 'اليوم 5',
  'اليوم 6', 'اليوم 7', 'اليوم 8', 'اليوم 9', 'اليوم 10',
  'الأساسي اليومي', 'البونص اليومي', 'الحالة', 'آخر تحديث',
  'الملاحظات', 'النقاط المعدلة', 'العداد', 'الأهداف الشخصية'
];

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function safeParse_(value, fallback) {
  try {
    if (!value || value == '') return fallback;
    return JSON.parse(value);
  } catch (err) {
    return fallback;
  }
}

function findRowByName_(values, name) {
  if (!name) return -1;
  var trimmedName = String(name).trim();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === trimmedName) return i + 1;
  }
  return -1;
}

function emptyRow_(name, notes) {
  return [
    name, 0, '—', 0,
    0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
    '[]', '[]',
    '{}', new Date().toLocaleString('ar'), notes || '{}', '{}', 0, '{}'
  ];
}

function doGet(e) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var type = e && e.parameter ? e.parameter.type : null;

  if (type === 'deleteGirl' || type === 'resetGirl') {
    var name = e.parameter.name;
    var rowNum = findRowByName_(values, name);
    if (rowNum >= 2) {
      if (type === 'deleteGirl') {
        sheet.deleteRow(rowNum);
      } else {
        var oldNotes = values[rowNum - 1][18] || '{}';
        var row = emptyRow_(name, oldNotes);
        sheet.getRange(rowNum, 1, 1, row.length).setValues([row]);
      }
      return ContentService.createTextOutput('ok')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService.createTextOutput('error: girl not found')
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var girls = [];
  for (var i = 0; i < values.length; i++) {
    if (!values[i][0] || values[i][0] === 'الاسم') continue;
    var row = values[i];
    var dt = [];
    for (var j = 4; j < 14; j++) dt.push(Number(row[j]) || 0);
    girls.push({
      name: row[0],
      grand: Number(row[1]) || 0,
      avg: row[2] || '—',
      filled: row[3] || 0,
      dayTotals: dt,
      dayBase: safeParse_(row[14], []),
      dayBonus: safeParse_(row[15], []),
      state: safeParse_(row[16], {}),
      lastUpdate: row[17] || '',
      notes: safeParse_(row[18], {}),
      editedPts: safeParse_(row[19], {}),
      counter: Number(row[20]) || 0,
      personalGoals: safeParse_(row[21], {})
    });
  }

  var jsonResult = JSON.stringify({ girls: girls });
  var callback = e && e.parameter ? e.parameter.callback : null;
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$.]{0,99}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + jsonResult + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonResult)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = getSheet_();
  var raw = '';
  try { raw = e.parameter.payload || e.postData.contents || ''; } catch (err) {}
  var data = {};
  try { data = JSON.parse(raw); } catch (err) {}

  if (!data.name || !String(data.name).trim()) {
    return ContentService.createTextOutput('error: missing name');
  }

  var values = sheet.getDataRange().getValues();
  var rowIndex = findRowByName_(values, data.name);

  if (data.type === 'dayNote') {
    var notes = {};
    if (rowIndex > 1) notes = safeParse_(values[rowIndex - 1][18], {});
    notes[data.day] = { text: data.text, time: new Date().toLocaleString('ar') };
    if (rowIndex > 1) {
      sheet.getRange(rowIndex, 19).setValue(JSON.stringify(notes));
    } else {
      sheet.appendRow(emptyRow_(String(data.name).trim(), JSON.stringify(notes)));
    }
    return ContentService.createTextOutput('ok');
  }

  var dt = Array.isArray(data.dayTotals) ? data.dayTotals.slice(0, 10) : [];
  while (dt.length < 10) dt.push(0);

  var dayBase = Array.isArray(data.dayBase) ? data.dayBase.slice(0, 10) : dt.map(function () { return 0; });
  var dayBonus = Array.isArray(data.dayBonus) ? data.dayBonus.slice(0, 10) : dt.map(function () { return 0; });
  while (dayBase.length < 10) dayBase.push(0);
  while (dayBonus.length < 10) dayBonus.push(0);

  var oldNotesForSend = rowIndex > 1 ? (values[rowIndex - 1][18] || '{}') : '{}';
  var rowData = [
    String(data.name).trim(),
    Number(data.grand) || 0,
    data.avg || '—',
    data.filled || 0,
    dt[0], dt[1], dt[2], dt[3], dt[4],
    dt[5], dt[6], dt[7], dt[8], dt[9],
    JSON.stringify(dayBase),
    JSON.stringify(dayBonus),
    JSON.stringify(data.state || {}),
    new Date().toLocaleString('ar'),
    oldNotesForSend,
    JSON.stringify(data.editedPts || {}),
    Number(data.counter) || 0,
    JSON.stringify(data.personalGoals || {})
  ];

  if (rowIndex > 1) {
    sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return ContentService.createTextOutput('ok');
}
