// ============================================================
//  Smart Sow Management System — Code.gs
//  Backend: Google Apps Script + Google Sheets + Gemini API
//  Version: 2.0  |  Author: อ๊อด x Claude
// ============================================================

// ─── CONFIG ─────────────────────────────────────────────────
const CONFIG = {
  GEMINI_API_KEY: PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '',
  GEMINI_MODEL:   'gemini-1.5-flash-latest',
  GEMINI_URL:     'https://generativelanguage.googleapis.com/v1beta/models/',

  SHEETS: {
    REGISTRY:    'แม่_ทะเบียนประวัติ',
    MATING:      'แม่_บันทึกผสม',
    FARROWING:   'แม่_บันทึกคลอด',
    WEANING:     'แม่_บันทึกหย่านม',
    BOAR:        'แม่_ทะเบียนพ่อพันธุ์',
    MEDICATION:  'แม่_การใช้ยา',
  },

  // คอลัมน์ภาษาไทย → index (0-based)
  COLS: {
    REGISTRY:   { earTag:'เบอร์หู', breed:'สายพันธุ์', dob:'วันเกิด', status:'สถานะ', parity:'ครั้งที่คลอด', note:'หมายเหตุ' },
    MATING:     { earTag:'เบอร์หู', matingDate:'วันผสม', boarTag:'เบอร์พ่อ', method:'วิธีผสม', tech:'ผู้ผสม', note:'หมายเหตุ' },
    FARROWING:  { earTag:'เบอร์หู', farrowDate:'วันคลอด', liveBorn:'เกิดมีชีวิต', stillBorn:'ตายคลอด', mummified:'มัมมี่', totalWeight:'น้ำหนักรวม', note:'หมายเหตุ' },
    WEANING:    { earTag:'เบอร์หู', weanDate:'วันหย่านม', weanCount:'จำนวนลูกหย่านม', totalWeanWeight:'น้ำหนักรวมหย่านม', note:'หมายเหตุ' },
    MEDICATION: { earTag:'เบอร์หู', medDate:'วันที่ฉีด', medicine:'ชื่อยา', dose:'ขนาดยา', reason:'เหตุผล', vet:'ผู้ฉีด', note:'หมายเหตุ' },
    BOAR:       { earTag:'เบอร์หู', breed:'สายพันธุ์', dob:'วันเกิด', status:'สถานะ', note:'หมายเหตุ' },
  },

  // ระยะเวลา (วัน)
  GESTATION_DAYS:    114,
  LACTATION_DAYS:    21,
  EMPTY_AFTER_DAYS:  21,
};

// ─── UTILS ──────────────────────────────────────────────────
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`ไม่พบชีต: ${name}`);
  return sh;
}

function sheetToObjects(sheetName) {
  const sh = getSheet(sheetName);
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows
    .filter(r => r.some(c => c !== ''))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
}

function appendRow(sheetName, colMap, data) {
  const sh = getSheet(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    const key = Object.keys(colMap).find(k => colMap[k] === h);
    return key ? (data[key] ?? '') : '';
  });
  sh.appendRow(row);
}

function response(ok, data, msg = '') {
  return ContentService
    .createTextOutput(JSON.stringify({ ok, data, msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── SOW STATUS CALCULATOR ──────────────────────────────────
function calcSowStatus(earTag) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // หาการผสมล่าสุด
  const matings = sheetToObjects(CONFIG.SHEETS.MATING)
    .filter(r => String(r[CONFIG.COLS.MATING.earTag]) === String(earTag))
    .sort((a, b) => new Date(b[CONFIG.COLS.MATING.matingDate]) - new Date(a[CONFIG.COLS.MATING.matingDate]));

  const lastMating = matings[0];
  if (!lastMating) return { status: 'ว่าง', detail: 'ไม่มีประวัติผสม' };

  const matingDate = new Date(lastMating[CONFIG.COLS.MATING.matingDate]);

  // หาการคลอดหลังผสมนี้
  const farrowings = sheetToObjects(CONFIG.SHEETS.FARROWING)
    .filter(r => {
      const t = String(r[CONFIG.COLS.FARROWING.earTag]) === String(earTag);
      const fd = new Date(r[CONFIG.COLS.FARROWING.farrowDate]);
      return t && fd >= matingDate;
    })
    .sort((a, b) => new Date(b[CONFIG.COLS.FARROWING.farrowDate]) - new Date(a[CONFIG.COLS.FARROWING.farrowDate]));

  const lastFarrow = farrowings[0];
  if (!lastFarrow) {
    // ยังไม่คลอด → คำนวณว่าอุ้มท้องหรือผสมใหม่
    const expectedFarrow = new Date(matingDate);
    expectedFarrow.setDate(expectedFarrow.getDate() + CONFIG.GESTATION_DAYS);
    const daysSinceMating = Math.floor((today - matingDate) / 86400000);
    if (daysSinceMating < CONFIG.GESTATION_DAYS) {
      return {
        status: 'อุ้มท้อง',
        detail: `ผสมเมื่อ ${Utilities.formatDate(matingDate, 'Asia/Bangkok', 'dd/MM/yyyy')}`,
        expectedFarrow: Utilities.formatDate(expectedFarrow, 'Asia/Bangkok', 'dd/MM/yyyy'),
        daysLeft: CONFIG.GESTATION_DAYS - daysSinceMating,
      };
    }
    return { status: 'ผสมแล้ว', detail: `ผสมเมื่อ ${Utilities.formatDate(matingDate, 'Asia/Bangkok', 'dd/MM/yyyy')} (เกิน ${CONFIG.GESTATION_DAYS} วัน)` };
  }

  const farrowDate = new Date(lastFarrow[CONFIG.COLS.FARROWING.farrowDate]);

  // หาการหย่านมหลังคลอดนี้
  const weanings = sheetToObjects(CONFIG.SHEETS.WEANING)
    .filter(r => {
      const t = String(r[CONFIG.COLS.WEANING.earTag]) === String(earTag);
      const wd = new Date(r[CONFIG.COLS.WEANING.weanDate]);
      return t && wd >= farrowDate;
    });

  if (weanings.length === 0) {
    // ยังไม่หย่านม = เลี้ยงลูก
    const expectedWean = new Date(farrowDate);
    expectedWean.setDate(expectedWean.getDate() + CONFIG.LACTATION_DAYS);
    return {
      status: 'เลี้ยงลูก',
      detail: `คลอดเมื่อ ${Utilities.formatDate(farrowDate, 'Asia/Bangkok', 'dd/MM/yyyy')}`,
      expectedWean: Utilities.formatDate(expectedWean, 'Asia/Bangkok', 'dd/MM/yyyy'),
      daysLeft: Math.max(0, Math.floor((expectedWean - today) / 86400000)),
    };
  }

  // หย่านมแล้ว → ว่าง
  const lastWean = weanings.sort((a, b) => new Date(b[CONFIG.COLS.WEANING.weanDate]) - new Date(a[CONFIG.COLS.WEANING.weanDate]))[0];
  return {
    status: 'ว่าง',
    detail: `หย่านมเมื่อ ${Utilities.formatDate(new Date(lastWean[CONFIG.COLS.WEANING.weanDate]), 'Asia/Bangkok', 'dd/MM/yyyy')}`,
  };
}

// ─── DASHBOARD ──────────────────────────────────────────────
function sow_getDashboardData() {
  try {
    const sows = sheetToObjects(CONFIG.SHEETS.REGISTRY);
    const summary = { ว่าง: 0, ผสมแล้ว: 0, อุ้มท้อง: 0, เลี้ยงลูก: 0, รวม: 0 };
    const alerts = [];

    sows.forEach(sow => {
      const earTag = sow[CONFIG.COLS.REGISTRY.earTag];
      if (!earTag) return;
      summary.รวม++;

      const statusObj = calcSowStatus(earTag);
      const st = statusObj.status;
      if (summary[st] !== undefined) summary[st]++;
      else summary[st] = 1;

      // แจ้งเตือน
      if (st === 'อุ้มท้อง' && statusObj.daysLeft <= 7) {
        alerts.push({ earTag, msg: `ใกล้คลอด ${statusObj.daysLeft} วัน (${statusObj.expectedFarrow})`, type: 'danger' });
      }
      if (st === 'เลี้ยงลูก' && statusObj.daysLeft <= 3) {
        alerts.push({ earTag, msg: `ใกล้หย่านม ${statusObj.daysLeft} วัน (${statusObj.expectedWean})`, type: 'warning' });
      }
      if (st === 'ว่าง') {
        alerts.push({ earTag, msg: `พร้อมผสม: ${statusObj.detail}`, type: 'info' });
      }
    });

    return { ok: true, summary, alerts: alerts.slice(0, 20) };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── SOW SEARCH ─────────────────────────────────────────────
function sow_searchByEarTag(earTag) {
  try {
    const sows = sheetToObjects(CONFIG.SHEETS.REGISTRY);
    const sow = sows.find(s => String(s[CONFIG.COLS.REGISTRY.earTag]) === String(earTag));
    if (!sow) return { ok: false, msg: `ไม่พบเบอร์หู: ${earTag}` };

    const statusObj = calcSowStatus(earTag);

    // ประวัติผสม 5 ครั้งล่าสุด
    const matings = sheetToObjects(CONFIG.SHEETS.MATING)
      .filter(r => String(r[CONFIG.COLS.MATING.earTag]) === String(earTag))
      .sort((a, b) => new Date(b[CONFIG.COLS.MATING.matingDate]) - new Date(a[CONFIG.COLS.MATING.matingDate]))
      .slice(0, 5);

    // ประวัติคลอด
    const farrowings = sheetToObjects(CONFIG.SHEETS.FARROWING)
      .filter(r => String(r[CONFIG.COLS.FARROWING.earTag]) === String(earTag))
      .sort((a, b) => new Date(b[CONFIG.COLS.FARROWING.farrowDate]) - new Date(a[CONFIG.COLS.FARROWING.farrowDate]))
      .slice(0, 5);

    return {
      ok: true,
      sow: {
        earTag,
        breed:  sow[CONFIG.COLS.REGISTRY.breed],
        dob:    sow[CONFIG.COLS.REGISTRY.dob] ? Utilities.formatDate(new Date(sow[CONFIG.COLS.REGISTRY.dob]), 'Asia/Bangkok', 'dd/MM/yyyy') : '-',
        parity: sow[CONFIG.COLS.REGISTRY.parity] || 0,
        note:   sow[CONFIG.COLS.REGISTRY.note] || '',
      },
      statusObj,
      matings:    matings.map(m => ({
        date:   m[CONFIG.COLS.MATING.matingDate] ? Utilities.formatDate(new Date(m[CONFIG.COLS.MATING.matingDate]), 'Asia/Bangkok', 'dd/MM/yyyy') : '-',
        boar:   m[CONFIG.COLS.MATING.boarTag],
        method: m[CONFIG.COLS.MATING.method],
        tech:   m[CONFIG.COLS.MATING.tech],
      })),
      farrowings: farrowings.map(f => ({
        date:        f[CONFIG.COLS.FARROWING.farrowDate] ? Utilities.formatDate(new Date(f[CONFIG.COLS.FARROWING.farrowDate]), 'Asia/Bangkok', 'dd/MM/yyyy') : '-',
        live:        f[CONFIG.COLS.FARROWING.liveBorn],
        still:       f[CONFIG.COLS.FARROWING.stillBorn],
        mummified:   f[CONFIG.COLS.FARROWING.mummified],
        totalWeight: f[CONFIG.COLS.FARROWING.totalWeight],
      })),
    };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── ACTION: บันทึกผสม ───────────────────────────────────────
function sow_recordMating(data) {
  try {
    appendRow(CONFIG.SHEETS.MATING, CONFIG.COLS.MATING, data);
    return { ok: true, msg: `บันทึกการผสมพันธุ์ เบอร์หู ${data.earTag} เรียบร้อยแล้ว` };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── ACTION: บันทึกคลอด ─────────────────────────────────────
function sow_recordFarrowing(data) {
  try {
    appendRow(CONFIG.SHEETS.FARROWING, CONFIG.COLS.FARROWING, data);
    // อัปเดต parity
    const sh = getSheet(CONFIG.SHEETS.REGISTRY);
    const [headers, ...rows] = sh.getDataRange().getValues();
    const earTagCol = headers.indexOf(CONFIG.COLS.REGISTRY.earTag);
    const parityCol = headers.indexOf(CONFIG.COLS.REGISTRY.parity);
    rows.forEach((row, i) => {
      if (String(row[earTagCol]) === String(data.earTag)) {
        sh.getRange(i + 2, parityCol + 1).setValue((row[parityCol] || 0) + 1);
      }
    });
    return { ok: true, msg: `บันทึกการคลอด เบอร์หู ${data.earTag} เรียบร้อยแล้ว` };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── ACTION: บันทึกหย่านม ────────────────────────────────────
function sow_recordWeaning(data) {
  try {
    appendRow(CONFIG.SHEETS.WEANING, CONFIG.COLS.WEANING, data);
    return { ok: true, msg: `บันทึกการหย่านม เบอร์หู ${data.earTag} เรียบร้อยแล้ว` };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── ACTION: บันทึกการใช้ยา ─────────────────────────────────
function sow_recordMedication(data) {
  try {
    appendRow(CONFIG.SHEETS.MEDICATION, CONFIG.COLS.MEDICATION, data);
    return { ok: true, msg: `บันทึกการใช้ยา เบอร์หู ${data.earTag} เรียบร้อยแล้ว` };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── GET BOAR LIST ───────────────────────────────────────────
function sow_getBoarList() {
  try {
    const boars = sheetToObjects(CONFIG.SHEETS.BOAR)
      .filter(b => b[CONFIG.COLS.BOAR.status] !== 'ปลด')
      .map(b => ({ earTag: b[CONFIG.COLS.BOAR.earTag], breed: b[CONFIG.COLS.BOAR.breed] }));
    return { ok: true, boars };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── GEMINI AI ASSISTANT "อ๊อดแอด" ─────────────────────────
function sow_chatWithAI(userMessage) {
  try {
    const dashboard = sow_getDashboardData();
    const summary   = dashboard.ok ? dashboard.summary : {};

    const systemContext = `
คุณคือ "อ๊อดแอด" ผู้ช่วย AI ผู้เชี่ยวชาญด้านการจัดการฟาร์มสุกร พูดภาษาไทยอย่างเป็นกันเอง
ตอบกระชับ ชัดเจน ใช้ภาษาที่เกษตรกรเข้าใจง่าย

ข้อมูลฟาร์ม ณ ขณะนี้:
- แม่หมูรวม: ${summary.รวม || 0} ตัว
- สถานะว่าง: ${summary.ว่าง || 0} ตัว (พร้อมผสม)
- สถานะผสมแล้ว: ${summary.ผสมแล้ว || 0} ตัว
- สถานะอุ้มท้อง: ${summary.อุ้มท้อง || 0} ตัว
- สถานะเลี้ยงลูก: ${summary.เลี้ยงลูก || 0} ตัว

ช่วยตอบคำถามเกี่ยวกับฟาร์ม การจัดการแม่พันธุ์ โรคสุกร โภชนาการ และการดูแลสุขภาพสัตว์
`.trim();

    const apiKey = CONFIG.GEMINI_API_KEY;
    if (!apiKey) return { ok: false, reply: 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ในหน้า Script Properties' };

    const url  = `${CONFIG.GEMINI_URL}${CONFIG.GEMINI_MODEL}:generateContent?key=${apiKey}`;
    const body = {
      system_instruction: { parts: [{ text: systemContext }] },
      contents: [{ role: 'user', parts: [{ text: userMessage }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 512 },
    };

    const res  = UrlFetchApp.fetch(url, {
      method:      'POST',
      contentType: 'application/json',
      payload:     JSON.stringify(body),
      muteHttpExceptions: true,
    });

    const json = JSON.parse(res.getContentText());
    const reply = json?.candidates?.[0]?.content?.parts?.[0]?.text || 'ขออภัย ไม่สามารถตอบได้ตอนนี้';
    return { ok: true, reply };
  } catch (e) {
    return { ok: false, reply: `เกิดข้อผิดพลาด: ${e.message}` };
  }
}

// ─── LIST ALL SOWS ───────────────────────────────────────────
function sow_getAllSows() {
  try {
    const sows = sheetToObjects(CONFIG.SHEETS.REGISTRY).map(s => ({
      earTag: s[CONFIG.COLS.REGISTRY.earTag],
      breed:  s[CONFIG.COLS.REGISTRY.breed],
    }));
    return { ok: true, sows };
  } catch (e) {
    return { ok: false, msg: e.message };
  }
}

// ─── doGet: Web App Entry Point ──────────────────────────────
function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Smart Sow Management — Niphon Farm')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}
