import fs from "node:fs";
import path from "node:path";
import * as XLSX from "xlsx";

type CandidateRow = Record<string, unknown>;

type StageKey = "application" | "phoneScreen" | "interview" | "offer" | "hire" | "join";
type OpenStageKey = "application" | "phoneScreen" | "interview" | "offer";

const SHEET_CANDIDATE = "candidate_master";

const COLUMNS = {
  idCard: "รหัสบัตรประชาชน",
  thaiPrefix: "คำนำหน้าชื่อ",
  thaiFirstName: "ชื่อ (ไทย)",
  thaiLastName: "สกุล (ไทย)",
  engFirstName: "ชื่อ (อังกฤษ)",
  engLastName: "สกุล (อังกฤษ)",
  email: "email",
  candidateStatus: "Candidate Status",
  createdDate: "วันที่สร้าง/Date",
  shortlist: "Shortlist",
  interview1: "1st round interview",
  interview2: "2nd round interview",
  interviewFinal: "Final round interview",
  offer: "Offering เสนอผลประโยชน์",
  hire: "Hiring",
  onboarding: "Onboarding",
  source: "Channel",
  turndownReason: "Turndown Reason",
  turndownDate: "Turndown Date",
} as const;

function getArg(flag: string, fallback?: string): string | undefined {
  const idx = process.argv.indexOf(flag);
  if (idx >= 0 && process.argv[idx + 1]) return process.argv[idx + 1];
  return fallback;
}

function resolveInputPath(): string {
  const explicit = getArg("--in");
  if (explicit) return path.resolve(explicit);

  const cwd = process.cwd();
  const candidates = [
    "src/output/recruitment-tracking.xlsx",
    "src/output/recruitement-tracking.xlsx",
    "output/recruitment-tracking.xlsx",
    "output/recruitement-tracking.xlsx",
  ].map((p) => path.resolve(cwd, p));

  const found = candidates.find((p) => fs.existsSync(p));
  if (!found) {
    throw new Error(
      `ไม่พบไฟล์ input (.xlsx) กรุณาระบุด้วย --in เช่น bun run dashboard --in src/output/recruitment-tracking.xlsx`,
    );
  }
  return found;
}

function resolveOutputPath(inputPath: string): string {
  const explicit = getArg("--out");
  if (explicit) return path.resolve(explicit);

  const dir = path.dirname(inputPath);
  return path.resolve(dir, "recruitment-dashboard.html");
}

function safeStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeReason(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function parseExcelDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    const date = new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const text = safeStr(value);
  if (!text) return null;

  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const [, d, m, y] = slashMatch;
    const date = new Date(Number(y), Number(m) - 1, Number(d));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const shortMonthMatch = text.match(/^(\d{1,2})-([A-Za-z]{3})-(\d{2})$/);
  if (shortMonthMatch) {
    const monthMap: Record<string, number> = {
      jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5,
      jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11,
    };
    const [, d, mon, y] = shortMonthMatch;
    const yy = Number(y);
    const fullYear = yy >= 70 ? 1900 + yy : 2000 + yy;
    const month = monthMap[mon.toLowerCase()];
    if (month === undefined) return null;
    const date = new Date(fullYear, month, Number(d));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const fallback = new Date(text);
  return Number.isNaN(fallback.getTime()) ? null : fallback;
}

function startOfDay(date: Date): Date {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function diffDays(from: Date | null, to: Date | null): number | null {
  if (!from || !to) return null;
  const ms = startOfDay(to).getTime() - startOfDay(from).getTime();
  if (ms < 0) return null;
  return Math.round(ms / 86_400_000);
}

function avg(values: Array<number | null>): number {
  const valid = values.filter((v): v is number => typeof v === "number");
  if (!valid.length) return 0;
  return Math.round(valid.reduce((a, b) => a + b, 0) / valid.length);
}

function pct(part: number, total: number): number {
  if (total <= 0) return 0;
  return Math.round((part / total) * 100);
}

function fmtDate(date: Date | null): string {
  if (!date) return "-";
  const day = String(date.getDate()).padStart(2, "0");
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const mon = months[date.getMonth()];
  const year = String(date.getFullYear()).slice(-2);
  return `${day}-${mon}-${year}`;
}

function getDate(row: CandidateRow, key: string): Date | null {
  return parseExcelDate(row[key]);
}

function getName(row: CandidateRow): string {
  const thai = [safeStr(row[COLUMNS.thaiPrefix]), safeStr(row[COLUMNS.thaiFirstName]), safeStr(row[COLUMNS.thaiLastName])]
    .filter(Boolean)
    .join(" ");
  if (thai) return thai;

  const eng = [safeStr(row[COLUMNS.engFirstName]), safeStr(row[COLUMNS.engLastName])].filter(Boolean).join(" ");
  if (eng) return eng;

  const id = safeStr(row[COLUMNS.idCard]);
  if (id) return id;

  return safeStr(row[COLUMNS.email]) || "N/A";
}

function isMeaningfulRow(row: CandidateRow): boolean {
  const id = safeStr(row[COLUMNS.idCard]);
  const thai = safeStr(row[COLUMNS.thaiFirstName]) || safeStr(row[COLUMNS.thaiLastName]);
  const eng = safeStr(row[COLUMNS.engFirstName]) || safeStr(row[COLUMNS.engLastName]);
  const email = safeStr(row[COLUMNS.email]);
  return Boolean(id || thai || eng || email);
}

function anyInterviewDate(row: CandidateRow): Date | null {
  const dates = [
    getDate(row, COLUMNS.interview1),
    getDate(row, COLUMNS.interview2),
    getDate(row, COLUMNS.interviewFinal),
  ].filter((d): d is Date => Boolean(d));
  if (!dates.length) return null;
  return new Date(Math.min(...dates.map((d) => d.getTime())));
}

function hasInterview(row: CandidateRow): boolean {
  return Boolean(anyInterviewDate(row));
}

function isHired(row: CandidateRow): boolean {
  const hireDate = getDate(row, COLUMNS.hire);
  const joinDate = getDate(row, COLUMNS.onboarding);
  if (hireDate || joinDate) return true;

  const status = safeStr(row[COLUMNS.candidateStatus]).toLowerCase();
  return status.includes("hire") || status.includes("onboard");
}

function isDeclined(row: CandidateRow): boolean {
  const reason = safeStr(row[COLUMNS.turndownReason]);
  const turndownDate = getDate(row, COLUMNS.turndownDate);
  if (reason || turndownDate) return true;

  const status = safeStr(row[COLUMNS.candidateStatus]).toLowerCase();
  return status.includes("declin") || status.includes("reject") || status.includes("turndown");
}

function getStage(row: CandidateRow): StageKey {
  if (getDate(row, COLUMNS.onboarding)) return "join";
  if (isHired(row)) return "hire";
  if (getDate(row, COLUMNS.offer)) return "offer";
  if (hasInterview(row)) return "interview";
  if (getDate(row, COLUMNS.shortlist)) return "phoneScreen";
  return "application";
}

function getOpenStage(row: CandidateRow): OpenStageKey {
  if (getDate(row, COLUMNS.offer)) return "offer";
  if (hasInterview(row)) return "interview";
  if (getDate(row, COLUMNS.shortlist)) return "phoneScreen";
  return "application";
}

function toIsoOrNull(date: Date | null): string | null {
  return date ? date.toISOString() : null;
}

function toDateInputValue(date: Date | null): string {
  if (!date) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function renderHtml(data: any): string {
  const payload = JSON.stringify(data).replace(/</g, "\\u003c");
  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Recruitment Dashboard</title>
  <style>
    :root {
      --bg: #f4f4f4;
      --card: #ffffff;
      --ink: #1f1f1f;
      --muted: #6a6a6a;
      --brand: #f3c01b;
      --panel: #2d323a;
      --line: #d8d8d8;
      --ok: #31a66a;
      --warn: #e07a18;
      --danger: #ef4f60;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Trebuchet MS", "Segoe UI", sans-serif;
      color: var(--ink);
      background:
        radial-gradient(circle at 8% -8%, #ffe88a 0, transparent 30%),
        radial-gradient(circle at 92% 110%, #f6d452 0, transparent 28%),
        var(--bg);
    }
    .container {
      width: min(1360px, 96vw);
      margin: 20px auto;
      display: grid;
      gap: 14px;
    }
    .filters {
      border: 2px solid #c99800;
      padding-bottom: 10px;
    }
    .filter-row {
      display: grid;
      grid-template-columns: repeat(6, minmax(0, 1fr));
      gap: 10px;
      align-items: end;
      margin-top: 8px;
    }
    .input-control {
      display: grid;
      gap: 5px;
    }
    .input-control label {
      font-size: 12px;
      text-transform: uppercase;
      color: #595959;
      font-weight: 700;
      letter-spacing: 0.5px;
    }
    .input-control select,
    .input-control input[type="date"] {
      width: 100%;
      border: 1px solid #cfcfcf;
      background: #fff;
      height: 35px;
      padding: 0 10px;
      font-size: 14px;
    }
    .check-wrap {
      display: flex;
      align-items: center;
      gap: 8px;
      height: 35px;
      border: 1px solid #cfcfcf;
      background: #fff;
      padding: 0 10px;
      font-size: 13px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .btn {
      height: 35px;
      border: 1px solid #111;
      background: #1d1d1d;
      color: #fff;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.7px;
      cursor: pointer;
    }
    .btn:hover { background: #000; }
    .filter-summary {
      margin-top: 10px;
      font-size: 13px;
      color: #4f4f4f;
      font-weight: 700;
    }
    .kpis {
      display: grid;
      grid-template-columns: repeat(5, minmax(0, 1fr));
      gap: 10px;
      background: var(--brand);
      border: 2px solid #c99800;
      padding: 12px;
    }
    .kpi {
      background: var(--panel);
      color: #fff;
      border: 1px solid #111;
      text-align: center;
      padding: 8px 6px;
    }
    .kpi-value {
      font-size: 36px;
      line-height: 1.1;
      font-weight: 800;
      letter-spacing: 0.5px;
    }
    .kpi-label {
      margin-top: 3px;
      font-size: 15px;
      font-weight: 700;
      color: #ffda5f;
      text-transform: uppercase;
      letter-spacing: 0.8px;
    }
    .grid {
      display: grid;
      grid-template-columns: 1.35fr 1.1fr 0.95fr;
      gap: 12px;
      align-items: start;
    }
    .card {
      background: var(--card);
      border: 1px solid var(--line);
      padding: 12px;
    }
    .card-title {
      background: #000;
      color: #fff;
      padding: 6px 10px;
      text-transform: uppercase;
      font-weight: 800;
      letter-spacing: 0.8px;
      font-size: 13px;
      margin: -12px -12px 12px -12px;
    }
    .funnel-row {
      display: grid;
      grid-template-columns: 135px 1fr 50px;
      align-items: center;
      gap: 10px;
      margin: 10px 0;
    }
    .funnel-label { font-size: 13px; text-transform: uppercase; font-weight: 700; }
    .bar-wrap { background: #f0f0f0; border: 1px solid #ddd; height: 24px; }
    .bar {
      height: 100%;
      background: linear-gradient(90deg, #f1c40f 0%, #f39c12 100%);
      color: #111;
      font-size: 12px;
      font-weight: 700;
      text-align: right;
      padding: 3px 8px 0 8px;
      white-space: nowrap;
    }
    .clickable { cursor: pointer; }
    .clickable:hover { opacity: 0.88; }
    .active-filter .bar-wrap { border-color: #111; }
    .active-filter .funnel-label { color: #000; }
    .donut-wrap {
      display: grid;
      grid-template-columns: 230px 1fr;
      gap: 12px;
      align-items: center;
      min-height: 240px;
    }
    .donut {
      width: 220px;
      height: 220px;
      border-radius: 50%;
      margin: 0 auto;
      position: relative;
      border: 1px solid #ddd;
    }
    .donut::before {
      content: "";
      position: absolute;
      inset: 37px;
      background: #fff;
      border-radius: 50%;
      border: 1px solid #e8e8e8;
      z-index: 1;
    }
    .donut-total {
      position: absolute;
      inset: 0;
      display: grid;
      place-items: center;
      font-size: 42px;
      font-weight: 900;
      z-index: 2;
      color: #222;
    }
    .legend {
      margin: 0;
      padding: 0;
      list-style: none;
      display: grid;
      gap: 8px;
    }
    .legend li {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      font-weight: 700;
      text-transform: uppercase;
      color: #333;
    }
    .dot {
      width: 11px;
      height: 11px;
      border-radius: 50%;
      flex: 0 0 11px;
    }
    .open-stage-row {
      display: grid;
      grid-template-columns: 125px 1fr 40px;
      align-items: center;
      gap: 8px;
      margin: 12px 0;
    }
    .open-stage-row .bar {
      background: #8d8d8d;
      color: #fff;
      text-align: center;
      padding: 3px 0 0 0;
    }
    .table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    .table th, .table td {
      border-bottom: 1px solid #ececec;
      padding: 6px 5px;
      text-align: left;
    }
    .table th { font-size: 12px; text-transform: uppercase; color: #4a4a4a; }
    .table tbody tr.clickable:hover { background: #fff6cf; }
    .table tbody tr.active-filter { background: #ffe8a0; }
    .mini-bar {
      height: 10px;
      background: #f1f1f1;
      border-radius: 2px;
      overflow: hidden;
    }
    .mini-bar > span {
      display: block;
      height: 100%;
      background: linear-gradient(90deg, #f3c443, #f2b207);
    }
    .mini-bar.red > span {
      background: linear-gradient(90deg, #f36f7a, #f14958);
    }
    .side-stack { display: grid; gap: 10px; }
    .info-block {
      border: 1px solid var(--line);
      text-align: center;
      background: #fafafa;
    }
    .info-head {
      background: #000;
      color: #fff;
      padding: 6px;
      font-size: 12px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.7px;
    }
    .info-value {
      padding: 14px 8px;
      font-size: 27px;
      font-weight: 800;
      color: #2b2b2b;
      line-height: 1.15;
    }
    .info-value.small { font-size: 22px; }
    .foot {
      color: var(--muted);
      font-size: 12px;
      text-align: right;
    }
    @media (max-width: 1120px) {
      .filter-row { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); }
      .grid { grid-template-columns: 1fr; }
      .donut-wrap { grid-template-columns: 1fr; justify-items: center; }
    }
  </style>
</head>
<body>
  <main class="container">
    <section class="card filters">
      <h3 class="card-title">Interactive Filters</h3>
      <div class="filter-row">
        <div class="input-control">
          <label for="filter-source">Source</label>
          <select id="filter-source"></select>
        </div>
        <div class="input-control">
          <label for="filter-stage">Stage</label>
          <select id="filter-stage">
            <option value="all">All Stages</option>
            <option value="application">Application</option>
            <option value="phoneScreen">Phone Screen</option>
            <option value="interview">Interview</option>
            <option value="offer">Offer</option>
            <option value="hire">Hire</option>
            <option value="join">Join</option>
          </select>
        </div>
        <div class="input-control">
          <label for="filter-from">From Date</label>
          <input id="filter-from" type="date" />
        </div>
        <div class="input-control">
          <label for="filter-to">To Date</label>
          <input id="filter-to" type="date" />
        </div>
        <label class="check-wrap">
          <input id="filter-open-only" type="checkbox" />
          Open Only
        </label>
        <button class="btn" id="filter-reset" type="button">Reset Filter</button>
      </div>
      <div class="filter-summary" id="filter-summary"></div>
    </section>

    <section class="kpis">
      <article class="kpi"><div class="kpi-value" id="kpi-applications"></div><div class="kpi-label">Applications</div></article>
      <article class="kpi"><div class="kpi-value" id="kpi-days-hire"></div><div class="kpi-label">Days to Hire</div></article>
      <article class="kpi"><div class="kpi-value" id="kpi-days-fill"></div><div class="kpi-label">Days to Fill</div></article>
      <article class="kpi"><div class="kpi-value" id="kpi-offer"></div><div class="kpi-label">Offer %</div></article>
      <article class="kpi"><div class="kpi-value" id="kpi-open"></div><div class="kpi-label">Open Applications</div></article>
    </section>

    <section class="grid">
      <article class="card">
        <h3 class="card-title">Recruitment</h3>
        <div id="funnel"></div>
      </article>

      <article class="card">
        <h3 class="card-title">Days For Each Stage</h3>
        <div class="donut-wrap">
          <div class="donut" id="donut"><div class="donut-total" id="donut-total"></div></div>
          <ul class="legend" id="legend"></ul>
        </div>
      </article>

      <article class="card">
        <h3 class="card-title">Open Applications In Each Stage</h3>
        <div id="open-stages"></div>
      </article>

      <article class="card">
        <h3 class="card-title">Source</h3>
        <table class="table">
          <thead>
            <tr><th>Source</th><th>Appl</th><th>% of Appl</th><th>Offer %</th></tr>
          </thead>
          <tbody id="source-body"></tbody>
        </table>
      </article>

      <article class="card">
        <h3 class="card-title">Decline Reasons</h3>
        <table class="table">
          <thead>
            <tr><th>Reason</th><th>Appl</th><th>Declined %</th></tr>
          </thead>
          <tbody id="decline-body"></tbody>
        </table>
      </article>

      <article class="card">
        <div class="side-stack">
          <section class="info-block">
            <div class="info-head">Hired Candidate</div>
            <div class="info-value small" id="hired-name">-</div>
          </section>
          <section class="info-block">
            <div class="info-head">Hired Date</div>
            <div class="info-value" id="hired-date">-</div>
          </section>
          <section class="info-block">
            <div class="info-head">Joining Date</div>
            <div class="info-value" id="joining-date">-</div>
          </section>
        </div>
      </article>
    </section>
    <p class="foot" id="foot-note"></p>
  </main>

  <script>
    const data = ${payload};

    const stageLabels = {
      application: 'Application',
      phoneScreen: 'Phone Screen',
      interview: 'Interview',
      offer: 'Offer',
      hire: 'Hire',
      join: 'Join',
    };
    const stageColors = {
      application: '#3f7ae0',
      phoneScreen: '#f15353',
      interview: '#f4bf1b',
      offer: '#38b06a',
      hire: '#ff7a00',
      join: '#58bec5',
    };
    const openStageOrder = ['application', 'phoneScreen', 'interview', 'offer'];
    const stageOrder = ['application', 'phoneScreen', 'interview', 'offer', 'hire', 'join'];

    const state = {
      source: 'ALL',
      stage: 'all',
      dateFrom: '',
      dateTo: '',
      openOnly: false,
    };

    function setText(id, value) {
      const el = document.getElementById(id);
      if (el) el.textContent = String(value ?? "-");
    }

    function asDate(iso) {
      if (!iso) return null;
      const d = new Date(iso);
      return Number.isNaN(d.getTime()) ? null : d;
    }

    function startOfDay(d) {
      return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }

    function diffDays(from, to) {
      if (!from || !to) return null;
      const ms = startOfDay(to).getTime() - startOfDay(from).getTime();
      if (ms < 0) return null;
      return Math.round(ms / 86400000);
    }

    function avg(values) {
      const valid = values.filter((v) => typeof v === 'number');
      if (!valid.length) return 0;
      const sum = valid.reduce((a, b) => a + b, 0);
      return Math.round(sum / valid.length);
    }

    function pct(part, total) {
      if (total <= 0) return 0;
      return Math.round((part / total) * 100);
    }

    function fmtDate(iso) {
      const date = asDate(iso);
      if (!date) return '-';
      const day = String(date.getDate()).padStart(2, '0');
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const mon = months[date.getMonth()];
      const year = String(date.getFullYear()).slice(-2);
      return day + '-' + mon + '-' + year;
    }

    function getFilterBaseDate(row) {
      return asDate(row.createdDate)
        || asDate(row.shortlistDate)
        || asDate(row.interviewDate)
        || asDate(row.offerDate)
        || asDate(row.hireDate)
        || asDate(row.onboardingDate);
    }

    function applyFilters(records) {
      return records.filter((row) => {
        if (state.source !== 'ALL' && row.source !== state.source) return false;
        if (state.stage !== 'all' && row.stage !== state.stage) return false;
        if (state.openOnly && (row.isHired || row.isDeclined)) return false;

        const baseDate = getFilterBaseDate(row);
        if (state.dateFrom) {
          const from = new Date(state.dateFrom + 'T00:00:00');
          if (!baseDate || startOfDay(baseDate).getTime() < from.getTime()) return false;
        }
        if (state.dateTo) {
          const to = new Date(state.dateTo + 'T00:00:00');
          if (!baseDate || startOfDay(baseDate).getTime() > to.getTime()) return false;
        }

        return true;
      });
    }

    function computeMetrics(records) {
      const applications = records.length;
      const offeredRows = records.filter((r) => Boolean(r.offerDate));
      const hiredRows = records.filter((r) => r.isHired);
      const openRows = records.filter((r) => !r.isHired && !r.isDeclined);

      const funnelCount = {
        application: applications,
        phoneScreen: records.filter((r) => Boolean(r.shortlistDate)).length,
        interview: records.filter((r) => Boolean(r.interviewDate)).length,
        offer: offeredRows.length,
        hire: hiredRows.length,
      };

      const funnel = {
        application: pct(funnelCount.application, applications),
        phoneScreen: pct(funnelCount.phoneScreen, applications),
        interview: pct(funnelCount.interview, applications),
        offer: pct(funnelCount.offer, applications),
        hire: pct(funnelCount.hire, applications),
      };

      const openStages = {
        application: 0,
        phoneScreen: 0,
        interview: 0,
        offer: 0,
      };
      openRows.forEach((row) => {
        openStages[row.openStage] += 1;
      });

      const daysToHire = avg(hiredRows.map((row) => {
        const start = asDate(row.createdDate) || asDate(row.shortlistDate) || asDate(row.interviewDate);
        const end = asDate(row.hireDate) || asDate(row.onboardingDate);
        return diffDays(start, end);
      }));

      const daysToFill = avg(hiredRows.map((row) => {
        const start = asDate(row.createdDate) || asDate(row.shortlistDate);
        const end = asDate(row.onboardingDate) || asDate(row.hireDate);
        return diffDays(start, end);
      }));

      const stageDays = {
        application: avg(records.map((row) => diffDays(asDate(row.createdDate), asDate(row.shortlistDate)))),
        phoneScreen: avg(records.map((row) => diffDays(asDate(row.shortlistDate), asDate(row.interviewDate)))),
        interview: avg(records.map((row) => diffDays(asDate(row.interviewDate), asDate(row.offerDate)))),
        offer: avg(records.map((row) => diffDays(asDate(row.offerDate), asDate(row.hireDate)))),
        hire: avg(records.map((row) => diffDays(asDate(row.hireDate), asDate(row.onboardingDate)))),
        join: 0,
      };

      const sourceMap = new Map();
      records.forEach((row) => {
        const stat = sourceMap.get(row.source) || { applications: 0, offers: 0 };
        stat.applications += 1;
        if (row.offerDate) stat.offers += 1;
        sourceMap.set(row.source, stat);
      });
      const source = [...sourceMap.entries()]
        .map(([name, stat]) => ({
          name,
          applications: stat.applications,
          applPct: pct(stat.applications, applications),
          offerPct: pct(stat.offers, stat.applications),
        }))
        .sort((a, b) => b.applications - a.applications)
        .slice(0, 8);

      const declinedRows = records.filter((row) => row.isDeclined);
      const declineMap = new Map();
      declinedRows.forEach((row) => {
        const reason = row.declineReason || 'UNSPECIFIED';
        declineMap.set(reason, (declineMap.get(reason) || 0) + 1);
      });
      const declines = [...declineMap.entries()]
        .map(([reason, count]) => ({ reason, count, pct: pct(count, declinedRows.length) }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 8);

      const latestHire = hiredRows
        .map((row) => ({
          name: row.name,
          hiredDate: row.hireDate || row.onboardingDate,
          joiningDate: row.onboardingDate,
        }))
        .filter((row) => Boolean(row.hiredDate))
        .sort((a, b) => (asDate(b.hiredDate)?.getTime() || 0) - (asDate(a.hiredDate)?.getTime() || 0))[0];

      return {
        kpi: {
          applications,
          daysToHire,
          daysToFill,
          offerPct: pct(offeredRows.length, applications),
          openApplications: openRows.length,
        },
        funnel,
        funnelCount,
        stageDays,
        openStages,
        source,
        declines,
        latestHire: {
          name: latestHire?.name || '-',
          hiredDate: latestHire ? fmtDate(latestHire.hiredDate) : '-',
          joiningDate: latestHire ? fmtDate(latestHire.joiningDate) : '-',
        },
      };
    }

    function renderKpis(metrics) {
      setText("kpi-applications", metrics.kpi.applications.toLocaleString());
      setText("kpi-days-hire", metrics.kpi.daysToHire.toLocaleString());
      setText("kpi-days-fill", metrics.kpi.daysToFill.toLocaleString());
      setText("kpi-offer", metrics.kpi.offerPct + "%");
      setText("kpi-open", metrics.kpi.openApplications.toLocaleString());
    }

    function renderFunnel(metrics) {
      const root = document.getElementById("funnel");
      if (!root) return;
      root.innerHTML = "";
      stageOrder.slice(0, 5).forEach((key) => {
        const row = document.createElement("div");
        row.className = "funnel-row clickable" + (state.stage === key ? " active-filter" : "");
        const p = metrics.funnel[key] || 0;
        row.innerHTML =
          '<div class="funnel-label">' + stageLabels[key] + '</div>'
          + '<div class="bar-wrap"><div class="bar" style="width:' + Math.max(6, p) + '%">' + p + '%</div></div>'
          + '<div style="font-weight:700">' + (metrics.funnelCount[key] || 0) + '</div>';
        row.onclick = () => {
          state.stage = state.stage === key ? 'all' : key;
          syncFiltersToUI();
          renderAll();
        };
        root.appendChild(row);
      });
    }

    function renderDonut(metrics) {
      const donut = document.getElementById("donut");
      const legend = document.getElementById("legend");
      if (!donut || !legend) return;

      const values = stageOrder.map((key) => Math.max(0, Number(metrics.stageDays[key] || 0)));
      const total = values.reduce((a, b) => a + b, 0);
      setText("donut-total", total);

      if (total <= 0) {
        donut.style.background = "#efefef";
      } else {
        let cursor = 0;
        const segments = stageOrder.map((key, idx) => {
          const v = values[idx];
          const from = (cursor / total) * 100;
          cursor += v;
          const to = (cursor / total) * 100;
          return \`\${stageColors[key]} \${from}% \${to}%\`;
        });
        donut.style.background = \`conic-gradient(\${segments.join(",")})\`;
      }

      legend.innerHTML = "";
      stageOrder.forEach((key) => {
        const li = document.createElement("li");
        li.innerHTML = '<span class="dot" style="background:' + stageColors[key] + '"></span>'
          + stageLabels[key] + ' - ' + (metrics.stageDays[key] || 0) + 'd';
        legend.appendChild(li);
      });
    }

    function renderOpenStages(metrics) {
      const root = document.getElementById("open-stages");
      if (!root) return;
      const max = Math.max(...openStageOrder.map((k) => metrics.openStages[k] || 0), 1);
      root.innerHTML = "";
      openStageOrder.forEach((key) => {
        const count = metrics.openStages[key] || 0;
        const row = document.createElement("div");
        row.className = "open-stage-row clickable" + (state.stage === key && state.openOnly ? " active-filter" : "");
        row.innerHTML =
          '<div class="funnel-label">' + stageLabels[key] + '</div>'
          + '<div class="bar-wrap"><div class="bar" style="width:' + Math.max(8, Math.round((count / max) * 100)) + '%">' + count + '</div></div>'
          + '<div style="font-weight:700">' + count + '</div>';
        row.onclick = () => {
          state.openOnly = true;
          state.stage = key;
          syncFiltersToUI();
          renderAll();
        };
        root.appendChild(row);
      });
    }

    function renderSource(metrics) {
      const body = document.getElementById("source-body");
      if (!body) return;
      body.innerHTML = "";
      if (!metrics.source.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = '<td colspan="4">No data</td>';
        body.appendChild(tr);
        return;
      }
      metrics.source.forEach((row) => {
        const tr = document.createElement("tr");
        tr.className = "clickable" + (state.source === row.name ? " active-filter" : "");
        tr.innerHTML =
          '<td>' + row.name + '</td>'
          + '<td>' + row.applications + '</td>'
          + '<td><div style="display:grid;grid-template-columns:52px 1fr;gap:8px;align-items:center;">'
          + '<span>' + row.applPct + '%</span>'
          + '<div class="mini-bar"><span style="width:' + Math.max(1, row.applPct) + '%"></span></div>'
          + '</div></td>'
          + '<td>' + row.offerPct + '%</td>';
        tr.onclick = () => {
          state.source = state.source === row.name ? 'ALL' : row.name;
          syncFiltersToUI();
          renderAll();
        };
        body.appendChild(tr);
      });
    }

    function renderDeclines(metrics) {
      const body = document.getElementById("decline-body");
      if (!body) return;
      body.innerHTML = "";
      if (!metrics.declines.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = '<td colspan="3">No decline data</td>';
        body.appendChild(tr);
        return;
      }
      metrics.declines.forEach((row) => {
        const tr = document.createElement("tr");
        tr.innerHTML =
          '<td>' + row.reason + '</td>'
          + '<td>' + row.count + '</td>'
          + '<td><div style="display:grid;grid-template-columns:52px 1fr;gap:8px;align-items:center;">'
          + '<span>' + row.pct + '%</span>'
          + '<div class="mini-bar red"><span style="width:' + Math.max(1, row.pct) + '%"></span></div>'
          + '</div></td>';
        body.appendChild(tr);
      });
    }

    function renderHiredInfo(metrics) {
      setText("hired-name", metrics.latestHire.name || "-");
      setText("hired-date", metrics.latestHire.hiredDate || "-");
      setText("joining-date", metrics.latestHire.joiningDate || "-");
    }

    function renderFilterSummary(filteredCount) {
      const parts = [];
      if (state.source !== 'ALL') parts.push('Source=' + state.source);
      if (state.stage !== 'all') parts.push('Stage=' + stageLabels[state.stage]);
      if (state.openOnly) parts.push('OpenOnly=Yes');
      if (state.dateFrom) parts.push('From=' + state.dateFrom);
      if (state.dateTo) parts.push('To=' + state.dateTo);
      const active = parts.length ? ' | Filters: ' + parts.join(', ') : '';
      setText('filter-summary', 'Showing ' + filteredCount.toLocaleString() + ' / ' + data.records.length.toLocaleString() + ' candidates' + active);
    }

    function initControls() {
      const sourceEl = document.getElementById('filter-source');
      if (sourceEl) {
        sourceEl.innerHTML = '';
        const allOpt = document.createElement('option');
        allOpt.value = 'ALL';
        allOpt.textContent = 'All Sources';
        sourceEl.appendChild(allOpt);

        (data.sources || []).forEach((src) => {
          const opt = document.createElement('option');
          opt.value = src;
          opt.textContent = src;
          sourceEl.appendChild(opt);
        });
      }

      const fromEl = document.getElementById('filter-from');
      const toEl = document.getElementById('filter-to');
      if (fromEl && data.dateBounds?.min) fromEl.value = data.dateBounds.min;
      if (toEl && data.dateBounds?.max) toEl.value = data.dateBounds.max;
      state.dateFrom = data.dateBounds?.min || '';
      state.dateTo = data.dateBounds?.max || '';

      syncFiltersToUI();

      sourceEl?.addEventListener('change', (e) => {
        state.source = e.target.value || 'ALL';
        renderAll();
      });
      document.getElementById('filter-stage')?.addEventListener('change', (e) => {
        state.stage = e.target.value || 'all';
        renderAll();
      });
      fromEl?.addEventListener('change', (e) => {
        state.dateFrom = e.target.value || '';
        renderAll();
      });
      toEl?.addEventListener('change', (e) => {
        state.dateTo = e.target.value || '';
        renderAll();
      });
      document.getElementById('filter-open-only')?.addEventListener('change', (e) => {
        state.openOnly = Boolean(e.target.checked);
        renderAll();
      });
      document.getElementById('filter-reset')?.addEventListener('click', () => {
        state.source = 'ALL';
        state.stage = 'all';
        state.openOnly = false;
        state.dateFrom = data.dateBounds?.min || '';
        state.dateTo = data.dateBounds?.max || '';
        syncFiltersToUI();
        renderAll();
      });
    }

    function syncFiltersToUI() {
      const sourceEl = document.getElementById('filter-source');
      const stageEl = document.getElementById('filter-stage');
      const fromEl = document.getElementById('filter-from');
      const toEl = document.getElementById('filter-to');
      const openEl = document.getElementById('filter-open-only');
      if (sourceEl) sourceEl.value = state.source;
      if (stageEl) stageEl.value = state.stage;
      if (fromEl) fromEl.value = state.dateFrom;
      if (toEl) toEl.value = state.dateTo;
      if (openEl) openEl.checked = state.openOnly;
    }

    function renderFoot() {
      setText("foot-note", 'Generated at ' + data.generatedAt + ' from ' + data.inputFile + ' | Click charts/rows to filter');
    }

    function renderAll() {
      const filtered = applyFilters(data.records || []);
      const metrics = computeMetrics(filtered);
      renderKpis(metrics);
      renderFunnel(metrics);
      renderDonut(metrics);
      renderOpenStages(metrics);
      renderSource(metrics);
      renderDeclines(metrics);
      renderHiredInfo(metrics);
      renderFilterSummary(filtered.length);
      renderFoot();
    }

    initControls();
    renderAll();
  </script>
</body>
</html>`;
}

function main() {
  const inputPath = resolveInputPath();
  const outputPath = resolveOutputPath(inputPath);
  const workbook = XLSX.readFile(inputPath, { cellDates: true });
  const sheet = workbook.Sheets[SHEET_CANDIDATE] || workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new Error(`ไม่พบชีตข้อมูลในไฟล์ ${inputPath}`);

  const rawRows = XLSX.utils.sheet_to_json<CandidateRow>(sheet, { defval: "" });
  const rows = rawRows.filter(isMeaningfulRow);

  const records = rows.map((row) => {
    const createdDate = getDate(row, COLUMNS.createdDate);
    const shortlistDate = getDate(row, COLUMNS.shortlist);
    const interviewDate = anyInterviewDate(row);
    const offerDate = getDate(row, COLUMNS.offer);
    const hireDate = getDate(row, COLUMNS.hire);
    const onboardingDate = getDate(row, COLUMNS.onboarding);
    const source = safeStr(row[COLUMNS.source]) || "UNKNOWN";

    return {
      name: getName(row),
      source: source.toUpperCase(),
      stage: getStage(row),
      openStage: getOpenStage(row),
      isHired: isHired(row),
      isDeclined: isDeclined(row),
      declineReason: normalizeReason(safeStr(row[COLUMNS.turndownReason])),
      createdDate: toIsoOrNull(createdDate),
      shortlistDate: toIsoOrNull(shortlistDate),
      interviewDate: toIsoOrNull(interviewDate),
      offerDate: toIsoOrNull(offerDate),
      hireDate: toIsoOrNull(hireDate),
      onboardingDate: toIsoOrNull(onboardingDate),
    };
  });

  const uniqueSources = Array.from(new Set(records.map((r) => r.source))).sort();
  const createdDates = records
    .map((r) => {
      const iso = r.createdDate || r.shortlistDate || r.interviewDate || r.offerDate || r.hireDate || r.onboardingDate;
      return iso ? new Date(iso) : null;
    })
    .filter((d): d is Date => Boolean(d) && !Number.isNaN(d!.getTime()))
    .sort((a, b) => a.getTime() - b.getTime());

  const data = {
    generatedAt: new Date().toLocaleString("th-TH"),
    inputFile: path.basename(inputPath),
    sources: uniqueSources,
    dateBounds: {
      min: toDateInputValue(createdDates[0] || null),
      max: toDateInputValue(createdDates[createdDates.length - 1] || null),
    },
    records,
  };

  const html = renderHtml(data);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, html, "utf8");

  console.log(`Dashboard generated: ${outputPath}`);
}

main();
