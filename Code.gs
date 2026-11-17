/**
 * Code.gs — ตัวกลางเก็บ/อ่านคะแนนแบบทดสอบ PISA 2029 (MAIL) ลง Google Sheet
 * ใช้ร่วมกันทุกสถานการณ์ (สถานการณ์ที่ 1-22) ไม่ต้องแก้ไฟล์นี้เมื่อเพิ่มสถานการณ์ใหม่
 *
 * หลักการทำงานของเวอร์ชันนี้
 *  1) ข้อปรนัย/จับคู่/เรียงลำดับ  → หน้าเว็บตรวจตามเกณฑ์ (รูบริค) แล้วส่งคะแนนมาเก็บ
 *     พร้อมคะแนนรายข้อในคอลัมน์ rubric_scores
 *  2) ข้ออัตนัย → ระบบ "ไม่ตรวจ" เก็บข้อความที่นักเรียนพิมพ์ไว้ในคอลัมน์ open_ended_answers
 *     สถานะเป็น "รอการตรวจคะแนน" ครูอ่านคำตอบแล้วพิมพ์คะแนนลงคอลัมน์ essay_score
 *     ระบบจะคำนวณ total_score / score_percent / สถานะ ให้เองทันที (ผ่าน onEdit)
 *  3) ข้อมูลกระบวนการ (เวลาที่ใช้ การเปิดดูหลักฐาน การดู/ฟังสื่อ) เก็บเหมือนเดิม
 *
 * ⚠️ โครงสร้างคอลัมน์ใหม่ไม่ตรงกับเวอร์ชันก่อนหน้า ถ้าพบแท็บ "responses" ที่หัวตารางเก่า
 *    ระบบจะเปลี่ยนชื่อแท็บเก่าเป็น "responses_old_<เวลา>" (ข้อมูลเดิมไม่หาย)
 *    แล้วสร้างแท็บ "responses" ใหม่ให้อัตโนมัติ
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 *  1. สร้าง Google Sheet ใหม่ → เมนู ส่วนขยาย > Apps Script
 *  2. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดทั้งไฟล์นี้แทน
 *  3. ปรับใช้ (Deploy) > การทำให้ใช้งานได้รายการใหม่ (New deployment) → Web app
 *     Execute as: Me / Who has access: Anyone
 *  4. คัดลอก URL (ลงท้าย /exec) ไปใส่ APPS_SCRIPT_URL ใน index.html และ scenario-XX.html ทุกไฟล์
 *
 * คอลัมน์ในแท็บ responses
 *  studentKey | name | grade | school | scenario_id | scenario_name
 *  | objective_score | objective_max      คะแนน/คะแนนเต็มของส่วนปรนัย (ตรวจโดยระบบตามเกณฑ์)
 *  | essay_max                            คะแนนเต็มของข้ออัตนัย (0 = ไม่มีข้ออัตนัย)
 *  | essay_status                         "รอการตรวจคะแนน" / "ตรวจแล้ว" / "ไม่มีข้ออัตนัย"
 *  | open_ended_answers                   คำตอบอัตนัยที่นักเรียนพิมพ์
 *  | essay_score                          ★ ครูกรอกคะแนนอัตนัยที่ช่องนี้
 *  | total_score | max_score
 *  | score_percent                        ร้อยละของส่วนที่ตรวจแล้ว (ถ้ายังรอตรวจ = ร้อยละของส่วนปรนัย)
 *  | rubric_scores                        คะแนนรายข้อตามเกณฑ์
 *  | time_spent_sec | evidence_opened_count | evidence_before_answer | media_watched_sec
 *  | details (JSON เต็ม สำรอง) | updated_at
 */

const SHEET_NAME = "responses";
const STATUS_PENDING = "รอการตรวจคะแนน";
const STATUS_GRADED = "ตรวจแล้ว";
const STATUS_NONE = "ไม่มีข้ออัตนัย";

const HEADER_ROW = [
  "studentKey", "name", "grade", "school",
  "scenario_id", "scenario_name",
  "objective_score", "objective_max", "essay_max", "essay_status",
  "open_ended_answers", "essay_score",
  "total_score", "max_score", "score_percent",
  "rubric_scores",
  "time_spent_sec", "evidence_opened_count", "evidence_before_answer", "media_watched_sec",
  "details", "updated_at",
];

// ดัชนีคอลัมน์ (เริ่มที่ 0) หาจากชื่อหัวตาราง ไม่ต้องไล่เลขเอง
const COL = {};
HEADER_ROW.forEach(function (name, i) { COL[name] = i; });

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (sheet && sheet.getLastRow() > 0) {
    const current = sheet.getRange(1, 1, 1, HEADER_ROW.length).getValues()[0];
    const same = HEADER_ROW.every(function (h, i) { return current[i] === h; });
    if (!same) {
      // หัวตารางเป็นเวอร์ชันเก่า → เก็บแท็บเดิมไว้ดูอ้างอิง แล้วเริ่มแท็บใหม่
      sheet.setName(SHEET_NAME + "_old_" + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyyMMdd_HHmmss"));
      sheet = null;
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER_ROW);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADER_ROW.length).setFontWeight("bold");
    // ช่องที่ครูต้องกรอก ทำสีเหลืองให้เห็นชัด
    sheet.getRange(1, COL.essay_score + 1).setBackground("#fff2cc");
  }
  return sheet;
}

/** แปลงเป็นตัวเลข ถ้าไม่ใช่ตัวเลขคืน null */
function toNum_(v) {
  if (v === "" || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

/**
 * คำนวณสถานะ/คะแนนรวม/ร้อยละ จากค่าดิบ
 * - ไม่มีข้ออัตนัย (essayMax = 0)        → status NONE, total = objective
 * - มีข้ออัตนัยแต่ครูยังไม่กรอกคะแนน     → status PENDING, total = objective (ยังไม่รวมอัตนัย)
 * - ครูกรอกคะแนน 0..essayMax              → status GRADED, total = objective + essay
 */
function computeTotals_(objectiveScore, objectiveMax, essayMax, essayScore) {
  const obj = toNum_(objectiveScore) || 0;
  const objMax = toNum_(objectiveMax) || 0;
  const eMax = toNum_(essayMax) || 0;
  const eScore = toNum_(essayScore);
  const maxScore = objMax + eMax;

  let status, total, basis;
  if (eMax <= 0) {
    status = STATUS_NONE; total = obj; basis = maxScore;
  } else if (eScore === null || eScore < 0 || eScore > eMax) {
    status = STATUS_PENDING; total = obj; basis = objMax;
  } else {
    status = STATUS_GRADED; total = obj + eScore; basis = maxScore;
  }
  return {
    status: status,
    total: total,
    maxScore: maxScore,
    percent: basis > 0 ? Math.round((total / basis) * 100) : "",
  };
}

/** คำนวณแถวใหม่หลังครูกรอก/แก้ essay_score */
function recalcRow_(sheet, row) {
  const r = sheet.getRange(row, 1, 1, HEADER_ROW.length).getValues()[0];
  const t = computeTotals_(r[COL.objective_score], r[COL.objective_max], r[COL.essay_max], r[COL.essay_score]);
  sheet.getRange(row, COL.essay_status + 1).setValue(t.status);
  sheet.getRange(row, COL.total_score + 1).setValue(t.total);
  sheet.getRange(row, COL.max_score + 1).setValue(t.maxScore);
  sheet.getRange(row, COL.score_percent + 1).setValue(t.percent);
}

/** Simple trigger: ทำงานเองทุกครั้งที่ครูแก้ช่องใน Sheet (ไม่ต้องตั้ง trigger) */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;
    const c1 = e.range.getColumn(), c2 = e.range.getLastColumn();
    const target = COL.essay_score + 1;
    if (target < c1 || target > c2) return; // ไม่ได้แก้คอลัมน์ essay_score
    const r1 = Math.max(e.range.getRow(), 2), r2 = e.range.getLastRow();
    for (let row = r1; row <= r2; row++) recalcRow_(sheet, row);
  } catch (err) { /* ไม่ให้ error รบกวนการแก้ไขชีต */ }
}

/** เรียกด้วยมือจากเมนู Run ได้ ถ้าต้องการคำนวณทุกแถวใหม่ (เช่น หลังวางข้อมูลจำนวนมากด้วยสคริปต์) */
function recalcAll() {
  const sheet = getOrCreateSheet_();
  for (let row = 2; row <= sheet.getLastRow(); row++) recalcRow_(sheet, row);
}

/**
 * ดึงตัวชี้วัดกระบวนการจาก details.behavior มาเป็นคอลัมน์แยก (เหมือนเวอร์ชันเดิม)
 * สถานการณ์ใดไม่มีตัวชี้วัดบางตัวจะเว้นว่าง ไม่ error
 */
function extractBehaviorColumns_(details) {
  const behavior = (details && details.behavior) || {};

  const timeSpentSec = behavior.time_spent_sec != null ? behavior.time_spent_sec : "";

  const evidenceOpenedCount = Array.isArray(behavior.evidence_opened)
    ? behavior.evidence_opened.length
    : "";

  let evidenceBeforeAnswer = "";
  if (behavior.evidence_before_answer != null) {
    evidenceBeforeAnswer = behavior.evidence_before_answer ? "ใช่" : "ไม่ใช่";
  }

  let mediaWatchedSec = "";
  if (behavior.video_watched_sec != null) mediaWatchedSec = behavior.video_watched_sec;
  else if (behavior.audio_listened_sec != null) mediaWatchedSec = behavior.audio_listened_sec;

  return {
    timeSpentSec: timeSpentSec,
    evidenceOpenedCount: evidenceOpenedCount,
    evidenceBeforeAnswer: evidenceBeforeAnswer,
    mediaWatchedSec: mediaWatchedSec,
  };
}

/**
 * เก็บคำตอบอัตนัยทุกข้อจาก details (ทุกข้อเก็บไว้ใต้คีย์ "text") รูปแบบ "Q1: ... | Q2: ..."
 * ข้ามกิ่ง behavior และ sim_log เพราะเป็นข้อมูลกระบวนการ ไม่ใช่คำตอบ
 */
function collectOpenEndedAnswers_(node, labelPath, out) {
  if (node === null || typeof node !== "object") return;

  if (Array.isArray(node)) {
    node.forEach(function (item) { collectOpenEndedAnswers_(item, labelPath, out); });
    return;
  }

  Object.keys(node).forEach(function (key) {
    if (key === "behavior" || key === "sim_log") return;

    const val = node[key];

    if (key === "text" && typeof val === "string" && val.trim() !== "") {
      out.push((labelPath || "คำตอบ") + ": " + val.trim());
      return;
    }

    if (val !== null && typeof val === "object") {
      const nextLabel = /^q[0-9]+$/i.test(key) ? key.toUpperCase() : labelPath;
      collectOpenEndedAnswers_(val, nextLabel, out);
    }
  });
}

/**
 * แปลง rubric (คะแนนรายข้อ) เป็นข้อความอ่านง่ายในชีต
 * rubric = [{ q, label, score, max, status }]  — ข้ออัตนัยจะมี status:"pending" และ score:null
 */
function formatRubric_(rubric) {
  if (!Array.isArray(rubric)) return "";
  return rubric.map(function (it) {
    const label = it.label || it.q || "";
    if (it.status === "pending" || it.score === null || it.score === undefined) {
      return label + ": " + STATUS_PENDING + " (เต็ม " + (it.max != null ? it.max : "?") + ")";
    }
    return label + ": " + it.score + "/" + it.max;
  }).join(" | ");
}

/**
 * POST — บันทึกผลสถานการณ์หนึ่ง (เขียนทับแถวเดิมถ้านักเรียนคนเดิมเคยทำสถานการณ์นี้)
 * body (JSON): { studentKey, name, grade, school, scenario_id, scenario_name,
 *                score,          // คะแนนส่วนปรนัย (ถ้าไม่มีอัตนัย = คะแนนรวม)
 *                max_score,      // คะแนนเต็มรวมปรนัย+อัตนัย
 *                objective_max,  // (ไม่บังคับ) คะแนนเต็มปรนัย
 *                essay_max,      // (ไม่บังคับ) คะแนนเต็มอัตนัย 0 ถ้าไม่มี
 *                rubric, details }
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet_();
    const values = sheet.getDataRange().getValues();

    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][COL.studentKey] === data.studentKey && values[i][COL.scenario_id] === data.scenario_id) {
        targetRow = i + 1;
        break;
      }
    }

    const maxScoreIn = toNum_(data.max_score) || 0;
    const essayMax = toNum_(data.essay_max) || 0;
    const objectiveMax = data.objective_max != null ? toNum_(data.objective_max) : (maxScoreIn - essayMax);
    const objectiveScore = toNum_(data.score) || 0;

    // ส่งซ้ำ = คำตอบใหม่ → ล้างคะแนนอัตนัยเดิมที่ครูเคยให้ เพื่อให้ตรวจใหม่
    const t = computeTotals_(objectiveScore, objectiveMax, essayMax, "");

    const b = extractBehaviorColumns_(data.details);
    const openEndedList = [];
    collectOpenEndedAnswers_(data.details, "", openEndedList);

    const row = new Array(HEADER_ROW.length).fill("");
    row[COL.studentKey] = data.studentKey || "";
    row[COL.name] = data.name || "";
    row[COL.grade] = data.grade || "";
    row[COL.school] = data.school || "";
    row[COL.scenario_id] = data.scenario_id || "";
    row[COL.scenario_name] = data.scenario_name || "";
    row[COL.objective_score] = objectiveScore;
    row[COL.objective_max] = objectiveMax;
    row[COL.essay_max] = essayMax;
    row[COL.essay_status] = t.status;
    row[COL.open_ended_answers] = openEndedList.join(" | ");
    row[COL.essay_score] = "";
    row[COL.total_score] = t.total;
    row[COL.max_score] = t.maxScore;
    row[COL.score_percent] = t.percent;
    row[COL.rubric_scores] = formatRubric_(data.rubric);
    row[COL.time_spent_sec] = b.timeSpentSec;
    row[COL.evidence_opened_count] = b.evidenceOpenedCount;
    row[COL.evidence_before_answer] = b.evidenceBeforeAnswer;
    row[COL.media_watched_sec] = b.mediaWatchedSec;
    row[COL.details] = JSON.stringify(data.details || {});
    row[COL.updated_at] = new Date().toISOString();

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        result: "success",
        essay_status: t.status,
        objective_score: objectiveScore,
        objective_max: objectiveMax,
        essay_max: essayMax,
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GET — ดึงความก้าวหน้าทั้งหมดของนักเรียนคนหนึ่ง  เรียก: <URL>?studentKey=xxxxx
 * คืน records[]: { scenario_id, scenario_name, objective_score, objective_max, essay_max,
 *                  essay_status ("pending"|"graded"|"none"), essay_score, total_score,
 *                  max_score, score_percent, score (= total_score เพื่อความเข้ากันได้), updated_at }
 * คำนวณจากคะแนนล่าสุดในชีตทุกครั้ง จึงเห็นคะแนนอัตนัยทันทีที่ครูกรอก
 */
function doGet(e) {
  try {
    const studentKey = e.parameter.studentKey;
    const sheet = getOrCreateSheet_();
    const values = sheet.getDataRange().getValues();
    const records = [];

    if (studentKey) {
      for (let i = 1; i < values.length; i++) {
        const r = values[i];
        if (r[COL.studentKey] !== studentKey) continue;

        const t = computeTotals_(r[COL.objective_score], r[COL.objective_max], r[COL.essay_max], r[COL.essay_score]);
        const code = t.status === STATUS_PENDING ? "pending" : (t.status === STATUS_GRADED ? "graded" : "none");

        records.push({
          scenario_id: r[COL.scenario_id],
          scenario_name: r[COL.scenario_name],
          objective_score: toNum_(r[COL.objective_score]) || 0,
          objective_max: toNum_(r[COL.objective_max]) || 0,
          essay_max: toNum_(r[COL.essay_max]) || 0,
          essay_status: code,
          essay_score: toNum_(r[COL.essay_score]),
          total_score: t.total,
          max_score: t.maxScore,
          score: t.total,
          score_percent: t.percent,
          updated_at: r[COL.updated_at],
        });
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", records: records }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
