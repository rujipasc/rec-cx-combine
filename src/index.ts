import fs from "node:fs";
import fsp from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import ExcelJS from "exceljs";
import * as XLSX from "xlsx";

// =========================
// CONFIG
// =========================
const SHEET_CANDIDATE = "candidate_master";
const SHEET_JR = "JR_Detail";

const CANDIDATE_KEY = "รหัสบัตรประชาชน";
const JR_KEY = "รหัสใบร้องขอ/ID";
const JR_NO_ALIAS = "JR No.";

// 10 columns base ที่ “อัปเดตได้” ตอน upsert
const CANDIDATE_BASE_COLS = [
  "คำนำหน้าชื่อ", "ชื่อ (ไทย)", "สกุล (ไทย)", "ชื่อ (อังกฤษ)",
  "สกุล (อังกฤษ)", "ชื่อเล่น", "รหัสบัตรประชาชน", "วันเกิด",
  "เบอร์ติดต่อ", "email",
] as const;

// Header order
const CANDIDATE_OUTPUT_HEADER_ORDER: string[] = [
  "คำนำหน้าชื่อ", "ชื่อ (ไทย)", "สกุล (ไทย)", "ชื่อ (อังกฤษ)",
  "สกุล (อังกฤษ)", "ชื่อเล่น", "รหัสบัตรประชาชน", "วันเกิด",
  "เบอร์ติดต่อ", "email", "JR No.", "หน่วยธุรกิจ/BU",
  "ตำแหน่งที่ขอรับ/Requested Position", "ประเภทการจ้าง/Employment Category",
  "ระดับ/Level", "ผู้รับผิดชอบ/Manage by", "สถานะล่าสุด/Latest status",
  "วันที่อัพเดทสถานะล่าสุด/Date Latest status", "สร้างโดย/Created by",
  "วันที่สร้าง/Date", "Candidate Status", "Shortlist", "1st round interview",
  "2nd round interview", "Final round interview", "Offering เสนอผลประโยชน์",
  "Hiring", "Onboarding", "Channel", "Turndown Reason", "Turndown Date",
  "Resume", "SLA by Level", "SLA (Shortlist)", "SLA (Interview)",
  "SLA (Offering)", "SLA (Hiring)", "SLA (Onboarding)",
];

const CANDIDATE_MAX_COLUMNS = 38;
const CANDIDATE_WIDTH = 24;
const HEADER_COLOR_A_J = "FF1CBBD8";
const HEADER_COLOR_K_V_AF = "FF5387D9";
const HEADER_COLOR_L_U_AG_AL = "FFFED243";

// =========================
// CLI ARGS
// =========================
function getArg(flag: string, fallback?: string) {
  const idx = process.argv.indexOf(flag);
  if (idx >= 0 && process.argv[idx + 1]) return process.argv[idx + 1];
  return fallback;
}

const CURRENT_DIR = process.cwd();
const IN_DIR = path.resolve(getArg("--in", path.join(CURRENT_DIR, "input"))!);
const OUT_FILE = path.resolve(getArg("--out", path.join(CURRENT_DIR, "output", "recruitment-tracking.xlsx"))!);

// Ensure output dir exists (Run Main Logic)
await mainWrapper();

// =========================
// UTIL
// =========================

async function ensureDir(p: string) {
  try {
    await fsp.mkdir(p, { recursive: true });
  } catch (err: any) {
    if (err.code === 'EEXIST') return;
    throw err;
  }
}

function normalizeHeader(h: any): string {
  if (h === null || h === undefined) return "";
  return String(h).replace(/\uFEFF/g, "").replace(/\s+/g, " ").trim();
}

function safeStr(v: any): string {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

async function fileExists(p: string) {
  try {
    await fsp.access(p, fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function acquireLock(lockPath: string, timeoutMs = 120_000, pollMs = 500) {
  const start = Date.now();
  while (true) {
    try {
      const fd = await fsp.open(lockPath, "wx");
      await fd.writeFile(`locked_at=${new Date().toISOString()}\nuser=${os.userInfo().username}\n`);
      await fd.close();
      return;
    } catch (e: any) {
      if (e?.code !== "EEXIST") throw e;
      if (Date.now() - start > timeoutMs) throw new Error(`Timeout waiting for lock: ${lockPath}`);
      await new Promise((r) => setTimeout(r, pollMs));
    }
  }
}

async function releaseLock(lockPath: string) {
  try {
    await fsp.unlink(lockPath);
  } catch { /* ignore */ }
}

function readWorkbook(filePath: string) {
  try {
    const fileBuffer = fs.readFileSync(filePath);
    return XLSX.read(fileBuffer, { type: "buffer", cellDates: true });
  } catch (error) {
    throw new Error(`Error reading file at ${filePath}: ${error}`);
  }
}

function sheetToObjects(wb: XLSX.WorkBook, sheetName: string, options?: { headerRow?: number }) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return [];
  const jsonOptions: XLSX.Sheet2JSONOpts = { defval: "" };
  if (options?.headerRow !== undefined) jsonOptions.range = options.headerRow;
  return XLSX.utils.sheet_to_json<Record<string, any>>(ws, jsonOptions);
}

function getSheetHeadersAtRow(ws: XLSX.WorkSheet, rowIndex: number): string[] {
  if (!ws) return [];
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  const headers: string[] = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const addr = XLSX.utils.encode_cell({ r: rowIndex, c });
    const cell = ws[addr];
    headers.push(normalizeHeader(cell?.v));
  }
  while (headers.length && !headers[headers.length - 1]) headers.pop();
  return headers.filter(Boolean);
}

function findHeaderRowInfo(wb: XLSX.WorkBook, sheetName: string, maxScanRows = 30) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return { headerRow: 0, headers: [] as string[], type: "UNKNOWN" as const };
  const range = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
  const start = range.s.r;
  const end = Math.min(range.e.r, start + maxScanRows - 1);
  const defaultHeaders = getSheetHeadersAtRow(ws, start);
  for (let row = start; row <= end; row++) {
    const headers = getSheetHeadersAtRow(ws, row);
    const type = detectTypeByHeaders(headers);
    if (type !== "UNKNOWN") return { headerRow: row, headers, type };
  }
  return { headerRow: start, headers: defaultHeaders, type: detectTypeByHeaders(defaultHeaders) };
}

function objectsToSheet(rows: Record<string, any>[], headerOrder?: string[]) {
  const allKeys = new Set<string>();
  for (const r of rows) Object.keys(r).forEach((k) => allKeys.add(k));
  const headers: string[] = [];
  if (headerOrder && headerOrder.length) for (const h of headerOrder) headers.push(h);
  for (const k of allKeys) if (!headers.includes(k)) headers.push(k);
  const normalized = rows.map((r) => {
    const out: Record<string, any> = {};
    for (const h of headers) out[h] = r[h] ?? "";
    return out;
  });
  return XLSX.utils.json_to_sheet(normalized, { header: headers });
}

function toCandidateBaseRow(row: Record<string, any>) {
  const out: Record<string, any> = {};
  for (const col of CANDIDATE_BASE_COLS) out[col] = row[col] ?? "";
  return out;
}

function buildHeaderOrder(rows: Record<string, any>[], preferred: string[] = []) {
  const allKeys = new Set<string>();
  for (const r of rows) Object.keys(r).forEach((k) => allKeys.add(k));
  const headers = [...preferred];
  for (const k of allKeys) if (!headers.includes(k)) headers.push(k);
  return headers;
}

function shiftFormulaRows(formula: string, rowDelta: number) {
  return formula.replace(
    /(?<![A-Z0-9_])(\$?[A-Z]{1,3})(\$?)(\d+)(?![A-Z0-9_])/gi,
    (_m, colRef: string, rowAbs: string, rowStr: string) => {
      const row = Number(rowStr);
      const next = rowAbs === "$" ? row : row + rowDelta;
      return `${colRef}${rowAbs}${next}`;
    },
  );
}

function cloneCellValue<T>(value: T): T {
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null as unknown as T;
    return new Date(value.getTime()) as T;
  }
  if (value && typeof value === "object") {
    try { return structuredClone(value); } catch { return JSON.parse(JSON.stringify(value)) as T; }
  }
  return value;
}

function excelCellValueToString(v: ExcelJS.CellValue): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") return safeStr(v);
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return ""; 
    return safeStr(v.toISOString());
  }
  if (typeof v === "object") {
    const anyV = v as any;
    if (typeof anyV.text === "string") return safeStr(anyV.text);
    if (typeof anyV.result === "string" || typeof anyV.result === "number") return safeStr(anyV.result);
    if (anyV.result instanceof Date) {
        if (isNaN(anyV.result.getTime())) return "";
        return safeStr(anyV.result.toISOString());
    }
    if (Array.isArray(anyV.richText)) return safeStr(anyV.richText.map((x: any) => x?.text ?? "").join(""));
  }
  return "";
}

function cloneCandidateRowTemplate(ws: ExcelJS.Worksheet, fromRow: number, toRow: number, maxCol: number) {
  for (let c = 1; c <= maxCol; c++) {
    const src = ws.getCell(fromRow, c);
    const dst = ws.getCell(toRow, c);

    if (src.style) dst.style = cloneCellValue(src.style);
    if (src.dataValidation) dst.dataValidation = cloneCellValue(src.dataValidation);

    if (c >= 22 && c <= 31) {
       dst.value = null;
    } else {
       const sv: any = src.value;
       if (sv && typeof sv === "object" && typeof sv.formula === "string") {
         dst.value = { formula: shiftFormulaRows(sv.formula, toRow - fromRow) };
       } else {
         dst.value = cloneCellValue(sv);
       }
    }
  }
}

// =========================================================
// [FINAL PERFECT v3] Safe Upsert + Fix SLA Date Format Issue
// =========================================================
function upsertCandidateSheetWithExcelJS(ws: ExcelJS.Worksheet, incoming: Record<string, any>[], headerOrder: string[]) {
  const maxCol = Math.min(CANDIDATE_MAX_COLUMNS, headerOrder.length);
  if (maxCol <= 0) return 1;

  for (let c = 1; c <= maxCol; c++) ws.getCell(1, c).value = headerOrder[c - 1] ?? "";
  if (ws.columnCount > maxCol) ws.spliceColumns(maxCol + 1, ws.columnCount - maxCol);

  const colByHeader = new Map<string, number>();
  for (let c = 1; c <= maxCol; c++) {
    const h = normalizeHeader(headerOrder[c - 1]);
    if (h) colByHeader.set(h, c);
  }

  const keyCol = colByHeader.get(CANDIDATE_KEY);
  if (!keyCol) throw new Error(`ไม่พบคอลัมน์ key "${CANDIDATE_KEY}" ใน ${SHEET_CANDIDATE}`);

  const baseColsInfo: Array<{ h: (typeof CANDIDATE_BASE_COLS)[number]; c: number }> = [];
  for (const h of CANDIDATE_BASE_COLS) {
    const c = colByHeader.get(h);
    if (c) baseColsInfo.push({ h, c });
  }

  const keyToRow = new Map<string, number>();
  let lastDataRow = 1; 
  const scanUntil = Math.max(ws.rowCount, ws.actualRowCount);
  
  for (let r = 2; r <= scanUntil; r++) {
    const key = excelCellValueToString(ws.getCell(r, keyCol).value);
    let hasBaseData = false;
    for (const b of baseColsInfo) {
      if (excelCellValueToString(ws.getCell(r, b.c).value)) hasBaseData = true;
    }
    if (key) keyToRow.set(key, r);
    if (key || hasBaseData) lastDataRow = r;
  }

  for (const row of incoming) {
    const key = safeStr(row[CANDIDATE_KEY]);
    if (!key) continue;

    let targetRow = keyToRow.get(key);
    
    // --- CASE 1: NEW ROW ---
    if (!targetRow) {
      targetRow = lastDataRow + 1;
      if (lastDataRow >= 2) cloneCandidateRowTemplate(ws, lastDataRow, targetRow, maxCol);
      
      for (let c = 22; c <= 31; c++) ws.getCell(targetRow, c).value = null;

      const incomingJR = row[JR_NO_ALIAS];
      if (incomingJR) ws.getCell(targetRow, 11).value = incomingJR;

      keyToRow.set(key, targetRow);
      lastDataRow = targetRow;
    }

    // --- CASE 2: EXISTING ROW ---
    const r = targetRow;

    // A. Base Data
    for (const b of baseColsInfo) {
        let val = row[b.h];
        if (val instanceof Date && isNaN(val.getTime())) val = "";
        ws.getCell(r, b.c).value = val ?? "";
    }

    // B. JR No. (Skipped for existing)

    // C. Update Formulas
    ws.getCell(r, 12).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,2,0),"")` };
    ws.getCell(r, 13).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,3,0),"")` };
    ws.getCell(r, 14).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,4,0),"")` };
    ws.getCell(r, 15).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,5,0),"")` };
    ws.getCell(r, 16).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,11,0),"")` };
    ws.getCell(r, 17).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,12,0),"")` };

    // Date Latest Status (R)
    const cellR = ws.getCell(r, 18);
    cellR.value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,13,0),"")` };
    cellR.numFmt = 'yyyy-mm-dd';

    ws.getCell(r, 19).value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,14,0),"")` };

    // Created Date (T)
    const cellT = ws.getCell(r, 20);
    cellT.value = { formula: `IFERROR(VLOOKUP($K${r},JR_Detail!$A:$O,15,0),"")` };
    cellT.numFmt = 'yyyy-mm-dd';

    // Candidate Status (U)
    ws.getCell(r, 21).value = { formula: `IF(AD${r}<>"", "Turndown", IFERROR(LOOKUP(2, 1/(V${r}:AB${r}<>""), V$1:AB$1), ""))` };

    // SLA Formulas
    // [FIXED] Force Format to "General" or "0" (Number) instead of Date
    const cellSLA_Level = ws.getCell(r, 33);
    cellSLA_Level.value = { formula: `IF(OR(ISNUMBER(SEARCH("Collector", M${r})), ISNUMBER(SEARCH("Underwriting", M${r})), ISNUMBER(SEARCH("Contact Center", M${r}))), 30, IF(OR(O${r}="Chief", O${r}="Head of"), 90, IF(OR(O${r}="Team lead", O${r}="Senior Professional", O${r}="Expert"), 60, IF(OR(O${r}="Professional", O${r}="Support"), 45, 0))))` };
    cellSLA_Level.numFmt = '0'; // Show as Integer (e.g. 30, 45, 60)

    const cellSLA_Shortlist = ws.getCell(r, 34);
    cellSLA_Shortlist.value = { formula: `IF(V${r}<>"", IFERROR(VALUE(V${r})-VALUE(T${r}), ""), "")` };
    cellSLA_Shortlist.numFmt = '0';

    const cellSLA_Interview = ws.getCell(r, 35);
    cellSLA_Interview.value = { formula: `IF(Y${r}<>"", IFERROR(VALUE(Y${r})-VALUE(T${r}), ""), IF(X${r}<>"", IFERROR(VALUE(X${r})-VALUE(T${r}), ""), IF(W${r}<>"", IFERROR(VALUE(W${r})-VALUE(T${r}), ""), "")))` };
    cellSLA_Interview.numFmt = '0';

    const cellSLA_Offering = ws.getCell(r, 36);
    cellSLA_Offering.value = { formula: `IF(Z${r}<>"", IFERROR(VALUE(Z${r})-VALUE(T${r}), ""), "")` };
    cellSLA_Offering.numFmt = '0';

    const cellSLA_Hiring = ws.getCell(r, 37);
    cellSLA_Hiring.value = { formula: `IF(AA${r}<>"", IFERROR(VALUE(AA${r})-VALUE(T${r}), ""), "")` };
    cellSLA_Hiring.numFmt = '0';

    const cellSLA_Onboarding = ws.getCell(r, 38);
    cellSLA_Onboarding.value = { formula: `IF(AB${r}<>"", IFERROR(VALUE(AB${r})-VALUE(T${r}), ""), "")` };
    cellSLA_Onboarding.numFmt = '0';

    // D. Data Validation
    for (let c = 22; c <= 28; c++) {
        const cell = ws.getCell(r, c);
        cell.dataValidation = {
            type: 'date', operator: 'greaterThan', showErrorMessage: true,
            errorTitle: 'Date Only', error: 'กรุณากรอกวันที่ (วว/ดด/ปปปป)', formulae: [new Date("2020-01-01")]
        };
        cell.numFmt = 'yyyy-mm-dd';
    }
    ws.getCell(r, 29).dataValidation = { type: 'list', allowBlank: true, formulae: ['"LinkedIn,JobThai,Referral,Agency,Walk-in"'] };
    ws.getCell(r, 30).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Salary,Counter Offer,Culture Fit,Ghosting,Skill Mismatch,Other"'] };
    
    const cellTurndownDate = ws.getCell(r, 31);
    cellTurndownDate.dataValidation = {
      type: 'date', operator: 'greaterThan', showErrorMessage: true,
      errorTitle: 'Date Only', error: 'กรุณากรอกวันที่ Turndown', formulae: [new Date("2020-01-01")]
    };
    cellTurndownDate.numFmt = 'yyyy-mm-dd';
  }

  return lastDataRow;
}

function applyCandidateFormattingWithExcelJS(ws: ExcelJS.Worksheet, lastRow: number) {
  const maxCol = Math.min(CANDIDATE_MAX_COLUMNS, ws.columnCount);
  const endRow = Math.max(1, lastRow);
  for (let c = 1; c <= maxCol; c++) ws.getColumn(c).width = CANDIDATE_WIDTH;
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } },
  };
  for (let r = 1; r <= endRow; r++) {
    for (let c = 1; c <= maxCol; c++) ws.getCell(r, c).border = thinBorder;
  }
  const fillHeader = (from: number, to: number, argb: string) => {
    const end = Math.min(to, maxCol);
    for (let c = from; c <= end; c++) ws.getCell(1, c).fill = { type: "pattern", pattern: "solid", fgColor: { argb } };
  };
  fillHeader(1, 10, HEADER_COLOR_A_J);
  fillHeader(11, 11, HEADER_COLOR_K_V_AF);
  fillHeader(22, 32, HEADER_COLOR_K_V_AF);
  fillHeader(12, 21, HEADER_COLOR_L_U_AG_AL);
  fillHeader(33, 38, HEADER_COLOR_L_U_AG_AL);

  ws.getRow(1).height = 32;
  const WHITE = "FFFFFFFF";
  const BLACK = "FF000000";

  function headerFontColorByCol(col: number) {
    if ((col >= 1 && col <= 10) || col === 11 || (col >= 22 && col <= 32)) return WHITE;
    return BLACK;
  }

  for (let c = 1; c <= maxCol; c++) {
    const cell = ws.getCell(1, c);
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.font = { ...(cell.font ?? {}), bold: true, size: 14, color: { argb: headerFontColorByCol(c) } };
  }
  
  for (let r = 1; r <= endRow; r++) ws.getCell(r, 1).font = { ...(ws.getCell(r, 1).font ?? {}), bold: true };
  ws.views = [{ state: "frozen", ySplit: 1 }];
  const lastColLetter = ws.getColumn(maxCol).letter;
  ws.autoFilter = { from: "A1", to: `${lastColLetter}1` };
}

function rewriteJRSheetWithExcelJS(wb: ExcelJS.Workbook, headers: string[], rows: Record<string, any>[]) {
  const old = wb.getWorksheet(SHEET_JR);
  if (old) wb.removeWorksheet(old.id);
  const ws = wb.addWorksheet(SHEET_JR);
  
  const cols = headers.length ? headers : ["_"];
  
  for (let c = 1; c <= cols.length; c++) ws.getCell(1, c).value = cols[c - 1];

  for (let i = 0; i < rows.length; i++) {
    for (let c = 1; c <= cols.length; c++) {
      const headerName = cols[c - 1];
      let val = rows[i][headerName];

      if (
        (headerName.includes("Date") || headerName.includes("วันที่")) && 
        typeof val === 'string' && 
        /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(val)
      ) {
        try {
          const parts = val.split('/');
          const dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          if (!isNaN(dateObj.getTime())) val = dateObj;
        } catch { /* ignore */ }
      }

      const cell = ws.getCell(i + 2, c);
      cell.value = val ?? "";
      if (val instanceof Date) cell.numFmt = 'dd/mm/yyyy';
    }
  }
}

function cleanHeaderForDetection(h: string) {
  return String(h || "").replace(/\uFEFF/g, "").replace(/[\u200B-\u200D\uFEFF]/g, "").replace(/\s+/g, "").trim();
}

function detectTypeByHeaders(headers: string[]) {
  const cleaned = headers.map(cleanHeaderForDetection);
  const hs = new Set(cleaned);
  const isCandidate = hs.has("รหัสบัตรประชาชน");
  const isJR = hs.has("รหัสใบร้องขอ/ID") || hs.has("JRNo.") || hs.has("JRNo");
  if (isCandidate && isJR) return "BOTH";
  if (isCandidate) return "CANDIDATE";
  if (isJR) return "JR";
  return "UNKNOWN";
}

async function listExcelFiles(dir: string) {
  const items = await fsp.readdir(dir, { withFileTypes: true });
  const files: string[] = [];
  for (const it of items) {
    if (!it.isFile()) continue;
    const ext = path.extname(it.name).toLowerCase();
    if (ext === ".xlsx" || ext === ".xls") files.push(path.join(dir, it.name));
  }
  return files;
}

async function pickLatestByType(inDir: string) {
  const files = await listExcelFiles(inDir);
  const candidates: { file: string; mtime: number }[] = [];
  const jrs: { file: string; mtime: number }[] = [];
  for (const f of files) {
    try {
      const stat = await fsp.stat(f);
      const wb = readWorkbook(f);
      const firstSheet = wb.SheetNames[0];
      const { type: t, headers } = findHeaderRowInfo(wb, firstSheet);
      console.log(`Checking file: ${path.basename(f)}`);
      console.log(` - Detected Type: ${t}`);
      console.log(` - Raw Headers (sample): ${headers.slice(0, 5).map(h => `'${h}'`).join(", ")}`);
      if (t === "CANDIDATE") candidates.push({ file: f, mtime: stat.mtimeMs });
      else if (t === "JR") jrs.push({ file: f, mtime: stat.mtimeMs });
      else if (t === "BOTH") {
        const name = path.basename(f).toLowerCase();
        if (name.includes("candidate")) candidates.push({ file: f, mtime: stat.mtimeMs });
        if (name.includes("jr")) jrs.push({ file: f, mtime: stat.mtimeMs });
      }
    } catch (e) {
      console.log(`Warning: Failed to read file ${path.basename(f)} - ${e}`);
    }
  }
  candidates.sort((a, b) => b.mtime - a.mtime);
  jrs.sort((a, b) => b.mtime - a.mtime);
  return { candidateFile: candidates[0]?.file, jrFile: jrs[0]?.file };
}

function readAllRowsFromFirstSheet(filePath: string) {
  const wb = readWorkbook(filePath);
  const sheetName = wb.SheetNames[0];
  const { headerRow } = findHeaderRowInfo(wb, sheetName);
  const rows = sheetToObjects(wb, sheetName, { headerRow });
  return { wb, sheetName, rows, headerRow };
}

async function mainWrapper() {
    try {
        await main();
    } catch (err: any) {
        console.error("❌ ERROR:", err?.message || err);
        console.log("Press any key to exit...");
        process.stdin.setRawMode(true);
        process.stdin.resume();
        process.stdin.on('data', process.exit.bind(process, 1));
    }
}

async function main() {
  console.log(`Working Directory: ${CURRENT_DIR}`);
  console.log(`Input Directory: ${IN_DIR}`);
  console.log(`Output File: ${OUT_FILE}`);

  await ensureDir(IN_DIR);
  await ensureDir(path.dirname(OUT_FILE));

  const { candidateFile, jrFile } = await pickLatestByType(IN_DIR);

  if (!candidateFile) throw new Error(`ไม่พบไฟล์ Candidate (.xlsx/.xls) ที่มีคอลัมน์ "${CANDIDATE_KEY}" ในโฟลเดอร์ ${IN_DIR}`);
  if (!jrFile) throw new Error(`ไม่พบไฟล์ JR (.xlsx/.xls) ที่มีคอลัมน์ "${JR_KEY}" หรือ "${JR_NO_ALIAS}" ในโฟลเดอร์ ${IN_DIR}`);

  const candParsed = readAllRowsFromFirstSheet(candidateFile);
  const jrParsed = readAllRowsFromFirstSheet(jrFile);
  const candIncoming = candParsed.rows.map(toCandidateBaseRow);
  const jrIncoming = jrParsed.rows;

  console.log("✅ Auto-detect input:");
  console.log(" - Candidate:", candidateFile);
  console.log(" - JR       :", jrFile);
  console.log(` - Candidate header row: ${candParsed.headerRow + 1}`);
  console.log(` - JR header row       : ${jrParsed.headerRow + 1}`);

  for (const r of jrIncoming) {
    if (r[JR_KEY] !== undefined && r[JR_NO_ALIAS] === undefined) {
      r[JR_NO_ALIAS] = r[JR_KEY];
    }
  }

  if (await fileExists(OUT_FILE)) {
    try { const outWb = readWorkbook(OUT_FILE); } catch (e) {
      throw new Error(`CRITICAL: Cannot read existing output file. Aborting.\nDetails: ${e}`);
    }
  }

  const jrFinal = jrIncoming;
  const jrHeaderOrder = buildHeaderOrder(jrFinal);
  
  const candidateHeaderOrder = [...CANDIDATE_OUTPUT_HEADER_ORDER].slice(0, CANDIDATE_MAX_COLUMNS);

  const lockPath = OUT_FILE + ".lock";

  await acquireLock(lockPath);
  try {
    const wb = new ExcelJS.Workbook();
    const outFileAlreadyExists = await fileExists(OUT_FILE);
    
    if (outFileAlreadyExists) {
      await wb.xlsx.readFile(OUT_FILE);
    }

    let wsCandidate = wb.getWorksheet(SHEET_CANDIDATE);
    if (!wsCandidate) wsCandidate = wb.addWorksheet(SHEET_CANDIDATE);
    
    const finalCandidateRows = upsertCandidateSheetWithExcelJS(
      wsCandidate,
      candIncoming,
      candidateHeaderOrder,
    );
    applyCandidateFormattingWithExcelJS(wsCandidate, finalCandidateRows);

    rewriteJRSheetWithExcelJS(wb, jrHeaderOrder, jrFinal);

    await wb.xlsx.writeFile(OUT_FILE);

    console.log("🎉 DONE");
    console.log(" - Fact file:", OUT_FILE);
    console.log(` - Sheets: ${SHEET_CANDIDATE}, ${SHEET_JR}`);
    console.log(` - Candidate rows: ${finalCandidateRows}`);
    console.log(` - JR rows: ${jrFinal.length}`);
  } finally {
    await releaseLock(lockPath);
  }
}