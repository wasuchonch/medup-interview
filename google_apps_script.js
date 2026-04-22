// ════════════════════════════════════════════════════════════════
//  MED UP Interview System — Google Apps Script
//  วิธีติดตั้ง:
//  1. เปิด Google Sheets ใหม่
//  2. Extensions → Apps Script
//  3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้
//  4. แก้ SHEET_ID บรรทัดด้านล่าง
//  5. Deploy → New deployment → Web app
//     Execute as: Me  |  Who has access: Anyone
//  6. Copy Web App URL → วางใน Dashboard แล้วกด "เชื่อม"
// ════════════════════════════════════════════════════════════════

var SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// ────────────────────────────────────────────────────────────────
//  doPost — รับข้อมูลคะแนนจาก Dashboard และบันทึกลง Sheet
// ────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var data = JSON.parse(e.postData.contents);

    // ── Sheet ของกรรมการแต่ละคน ──────────────────────────────
    var mgrName = (data.interviewer || 'Unknown').replace(/[\/\:\?\*\[\]]/g, '_');
    var ivSheet = ss.getSheetByName(mgrName);

    if (!ivSheet) {
      ivSheet = ss.insertSheet(mgrName);
      var h1 = [
        'วันที่', 'ชื่อผู้สมัคร', 'ตำแหน่ง',
        'M', 'E', 'D', 'U', 'P', 'C',
        'รวม', '%', 'ผล',
        'หมายเหตุ M', 'หมายเหตุ E', 'หมายเหตุ D',
        'หมายเหตุ U', 'หมายเหตุ P', 'หมายเหตุ C',
        'candId', 'mgrId'
      ];
      ivSheet.appendRow(h1);
      ivSheet.getRange(1, 1, 1, h1.length)
             .setFontWeight('bold')
             .setBackground('#0369a1')
             .setFontColor('#ffffff');
      ivSheet.setFrozenRows(1);
    }

    // ── ตรวจสอบว่ามี record เดิมหรือไม่ (upsert) ────────────
    var existingRows = ivSheet.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < existingRows.length; i++) {
      if (String(existingRows[i][18]) === String(data.candId) &&
          String(existingRows[i][19]) === String(data.mgrId)) {
        rowIndex = i + 1; // 1-based
        break;
      }
    }

    var row = [
      data.interviewDate  || '',
      data.candidateName  || '',
      data.position       || '',
      Number(data.scores && data.scores.M) || 0,
      Number(data.scores && data.scores.E) || 0,
      Number(data.scores && data.scores.D) || 0,
      Number(data.scores && data.scores.U) || 0,
      Number(data.scores && data.scores.P) || 0,
      Number(data.scores && data.scores.C) || 0,
      Number(data.total)  || 0,
      (Number(data.pct)   || 0) + '%',
      data.result         || '',
      (data.remarks && data.remarks.M) || '',
      (data.remarks && data.remarks.E) || '',
      (data.remarks && data.remarks.D) || '',
      (data.remarks && data.remarks.U) || '',
      (data.remarks && data.remarks.P) || '',
      (data.remarks && data.remarks.C) || '',
      data.candId || '',
      data.mgrId  || ''
    ];

    if (rowIndex > 0) {
      ivSheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    } else {
      ivSheet.appendRow(row);
    }

    // ── Sheet Summary (คะแนนรวมทุกคน) ───────────────────────
    var sumSheet = ss.getSheetByName('Summary');
    if (!sumSheet) {
      sumSheet = ss.insertSheet('Summary');
      ss.setActiveSheet(sumSheet);
      ss.moveActiveSheet(1);
      var h2 = [
        'วันที่', 'ชื่อผู้สมัคร', 'ตำแหน่ง', 'กรรมการ',
        'M', 'E', 'D', 'U', 'P', 'C',
        'รวม', '%', 'ผล',
        'candId', 'mgrId'
      ];
      sumSheet.appendRow(h2);
      sumSheet.getRange(1, 1, 1, h2.length)
              .setFontWeight('bold')
              .setBackground('#0c4a6e')
              .setFontColor('#ffffff');
      sumSheet.setFrozenRows(1);
    }

    var s2Rows   = sumSheet.getDataRange().getValues();
    var sumIndex = -1;
    for (var j = 1; j < s2Rows.length; j++) {
      if (String(s2Rows[j][13]) === String(data.candId) &&
          String(s2Rows[j][14]) === String(data.mgrId)) {
        sumIndex = j + 1;
        break;
      }
    }

    var srow = [
      data.interviewDate || '',
      data.candidateName || '',
      data.position      || '',
      data.interviewer   || '',
      Number(data.scores && data.scores.M) || 0,
      Number(data.scores && data.scores.E) || 0,
      Number(data.scores && data.scores.D) || 0,
      Number(data.scores && data.scores.U) || 0,
      Number(data.scores && data.scores.P) || 0,
      Number(data.scores && data.scores.C) || 0,
      Number(data.total) || 0,
      (Number(data.pct)  || 0) + '%',
      data.result        || '',
      data.candId        || '',
      data.mgrId         || ''
    ];

    if (sumIndex > 0) {
      sumSheet.getRange(sumIndex, 1, 1, srow.length).setValues([srow]);
    } else {
      sumSheet.appendRow(srow);
    }

    // ── ส่ง response กลับ ────────────────────────────────────
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', message: 'Saved successfully' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ────────────────────────────────────────────────────────────────
//  doGet — ส่งข้อมูลทั้งหมดกลับไปให้ Dashboard (CORS safe)
// ────────────────────────────────────────────────────────────────
function doGet(e) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Summary');

    if (!sheet || sheet.getLastRow() < 2) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rows    = sheet.getDataRange().getValues();
    var headers = rows[0];

    var records = rows.slice(1).map(function (r) {
      var o = {};
      headers.forEach(function (h, i) { o[h] = r[i]; });
      return {
        candId:        o['candId']     || '',
        mgrId:         o['mgrId']      || '',
        interviewDate: String(o['วันที่']       || ''),
        candidateName: String(o['ชื่อผู้สมัคร'] || ''),
        position:      String(o['ตำแหน่ง']      || ''),
        interviewer:   String(o['กรรมการ']       || ''),
        scores: {
          M: Number(o['M']) || 0,
          E: Number(o['E']) || 0,
          D: Number(o['D']) || 0,
          U: Number(o['U']) || 0,
          P: Number(o['P']) || 0,
          C: Number(o['C']) || 0
        },
        total:  Number(o['รวม'])         || 0,
        pct:    parseInt(o['%'])          || 0,
        result: String(o['ผล']            || '')
      };
    });

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ────────────────────────────────────────────────────────────────
//  ฟังก์ชันช่วยทดสอบ — รันใน Apps Script editor ได้เลย
// ────────────────────────────────────────────────────────────────
function testPost() {
  var mockData = {
    candId: 1, mgrId: 1,
    interviewDate: '2025-06-01',
    candidateName: 'นายทดสอบ ระบบ',
    position: 'อาจารย์แพทย์',
    interviewer: 'ผศ.นพ.สมชาย ใจดี',
    scores: { M:4, E:3, D:4, U:5, P:4, C:3 },
    remarks: { M:'', E:'', D:'', U:'', P:'', C:'' },
    total: 23, pct: 77, result: 'พิจารณาเพิ่มเติม'
  };
  var mock = { postData: { contents: JSON.stringify(mockData) } };
  var result = doPost(mock);
  Logger.log(result.getContent());
}

function testGet() {
  var result = doGet({});
  Logger.log(result.getContent());
}
