/** 
 * 🏛️ ENTERPRISE POLICE DOCUMENT SYSTEM
 * Version: 15.0 (Official Stable - PDF Fix & Auto-Fill Sync)
 * พัฒนาโดย: พี่ชาย สำหรับ กันต์
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const DATA_SHEET = 'DATA_RAW';
const MAP_SHEET = 'FIELD_MAP';
const PREFIX = 'TPL_';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบเอกสารดิจิทัล (กันต์)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 📂 1. ดึงรายชื่อ Template ทั้งหมด */
function getTemplates() {
  return SS.getSheets().filter(s => s.getName().startsWith(PREFIX)).map(s => s.getName());
}

/** 🛠️ 2. FIELD MANAGER: ดึงข้อมูลและเรียงตามพิกัด (Row/Col) จากชีทจริง */
function getAllMasterFields() {
  try {
    let sheet = SS.getSheetByName(MAP_SHEET) || createMapSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.map((r, i) => ({
      id: i + 2,
      field_name: String(r[0] || ''),
      template_name: String(r[1] || ''),
      row: Number(r[2] || 0),
      column: Number(r[3] || 0),
      type: String(r[4] || 'text'),
      description: String(r[5] || '')
    })).sort((a, b) => {
      if (a.row !== b.row) return a.row - b.row;
      return a.column - b.column;
    });
  } catch (e) { return []; }
}

/** 🔍 3. SMART SYNC: ระบบแชร์ฟิลด์ข้าม Template */
function syncTemplateFields(templateName) {
  try {
    const tplSheet = SS.getSheetByName(templateName);
    const mapSheet = SS.getSheetByName(MAP_SHEET) || createMapSheet();
    const data = tplSheet.getDataRange().getValues();
    const regex = /\[([a-zA-Z0-9_]+)\]/g;
    const allFields = getAllMasterFields();
    
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        let match;
        while ((match = regex.exec(String(data[r][c]))) !== null) {
          const fName = match[1];
          const globalMatch = allFields.find(f => f.field_name === fName);
          const localMatch = allFields.find(f => f.field_name === fName && f.template_name === templateName);
          if (!localMatch) {
            mapSheet.appendRow([fName, templateName, r + 1, c + 1, (globalMatch ? globalMatch.type : 'text'), (globalMatch ? globalMatch.description : fName)]);
          } else {
            mapSheet.getRange(localMatch.id, 3, 1, 2).setValues([[r + 1, c + 1]]);
          }
        }
      }
    }
    return { status: 'success' };
  } catch(e) { return { status: 'error', message: e.toString() }; }
}

/** 🔄 4. AUTO-FILL: ดึงข้อมูล JSON ล่าสุดจาก DATA_RAW */
function getLastEntryData() {
  try {
    const sheet = SS.getSheetByName(DATA_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return null;
    return JSON.parse(sheet.getRange(sheet.getLastRow(), 4).getValue());
  } catch (e) { return null; }
}

/** ➕ 5. บันทึกข้อมูลและสร้างชีทใหม่ */
function processFormSubmission(formData, templateName) {
  try {
    const dataSheet = SS.getSheetByName(DATA_SHEET) || SS.insertSheet(DATA_SHEET);
    dataSheet.appendRow([new Date(), templateName, formData.consent_name || 'N/A', JSON.stringify(formData)]);
    const timeStamp = Utilities.formatDate(new Date(), "GMT+7", "HHmm");
    const newName = `${formData.consent_name || 'บันทึก'}_${timeStamp}`;
    const newSheet = SS.getSheetByName(templateName).copyTo(SS).setName(newName);
    
    const data = newSheet.getDataRange().getValues();
    const regex = /\[([a-zA-Z0-9_]+)\]/g;
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        let cellText = String(data[r][c]);
        if (cellText.includes('[')) {
          const tokens = cellText.match(regex);
          if (tokens) {
            let updatedText = cellText;
            tokens.forEach(t => {
              const fName = t.replace('[','').replace(']','');
              updatedText = updatedText.replace(new RegExp(`\\[${fName}\\]`, 'g'), formData[fName] || "");
            });
            newSheet.getRange(r+1, c+1).setValue(updatedText).setFontFamily("TH SarabunPSK").setFontSize(16);
          }
        }
      }
    }
    SS.setActiveSheet(newSheet);
    return { status: 'success', sheetName: newName };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

/** 📄 6. PDF URL: สร้างลิงก์สำหรับการ Export แบบตั้งค่าหน้ากระดาษได้ */
function getPdfUrl(sheetName) {
  try {
    const sheet = SS.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'ไม่พบชีท' };
    const base = SS.getUrl().replace(/edit$/, '') + 'export?';
    const opts = 'exportFormat=pdf&format=pdf&size=A4&portrait=true&fitw=true&gid=' + sheet.getSheetId();
    return { status: 'success', url: base + opts };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

/** 📜 7. HISTORY: ดึงประวัติรายการ */
function getDocumentList() {
  const sheet = SS.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1).map((r, i) => ({
    id: i + 2, tpl: r[1], timestamp: Utilities.formatDate(new Date(r[0]), "GMT+7", "dd/MM/yyyy HH:mm"), name: r[2]
  })).reverse();
}

function deleteDocument(rowId, name) {
  const s = SS.getSheetByName(name); if (s) SS.deleteSheet(s);
  SS.getSheetByName(DATA_SHEET).deleteRow(rowId);
  return { status: 'success' };
}

function updateMasterField(id, data) {
  const mapSheet = SS.getSheetByName(MAP_SHEET);
  const rows = mapSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.field_name) mapSheet.getRange(i+1, 5, 1, 2).setValues([[data.type, data.description]]);
  }
  return { status: 'success' };
}

function deleteMasterField(id) { SS.getSheetByName(MAP_SHEET).deleteRow(id); return { status: 'success' }; }
function createMapSheet() { const s = SS.insertSheet(MAP_SHEET); s.appendRow(['field_name', 'template_name', 'row', 'column', 'type', 'description']); return s; }
/** 🇹🇭 ฟังก์ชันแปลงวันที่เป็นรูปแบบไทยย่อ (เช่น 2 ก.ย.69) */
function formatThaiDateShort(date) {
  const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  const d = new Date(date);
  const day = d.getDate();
  const month = months[d.getMonth()];
  const year = (d.getFullYear() + 543).toString().slice(-2); // เอาปี พ.ศ. 2 หลักท้าย
  return `${day} ${month}${year}`;
}






