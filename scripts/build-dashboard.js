const fs = require('fs');
const https = require('https');
const path = require('path');
const { createRequire } = require('module');
const ExcelJS = require('exceljs');
const { AGENDA_DAYS } = require('./agenda-data');

const workbookPath = process.env.SUMMIT_XLSX || path.resolve(__dirname, '..', 'assets', 'source', 'registration-report-deduped-2026-08-06.xlsx');
const outDir = path.resolve(__dirname, '..');
const dataDir = path.join(outDir, 'assets', 'data');
const workspaceDir = path.resolve(outDir, '..');
const workspaceRequire = createRequire(path.join(workspaceDir, 'gmail-tool', 'package.json'));
const XLSX = workspaceRequire('xlsx');
const API_VERSION = 'v59.0';
const SPONSOR_OPPORTUNITY_TYPE = 'Summit Sponsorship';

const GOALS = {
  totalRegistrants: 400,
  paidTickets: 165,
  ticketRevenue: 74000,
  sponsorRevenue: 313950,
  totalRevenue: 387950,
  sponsorPasses: 57,
  uniqueSponsors: 25,
  staffTickets: 60,
  customerProspectRegistrations: 283,
  customerProspectPaid: 165,
  customerProspectComped: 118,
  projectedCost: 503636,
  netTarget: -115686
};

const ACTUAL_2025 = {
  totalRegistrations: 301,
  staffTickets: 60,
  customerProspectRegistrations: 179,
  customerProspectPaid: 85,
  customerProspectComped: 94,
  avgTicketPrice: 471,
  ticketRevenue: 40040,
  sponsorTickets: 62,
  sponsors: 21,
  sponsorshipRevenue: 184000,
  totalRevenue: 224040,
  totalCost: 336402,
  totalCostTe: 315171,
  netProfit: -112362,
  profitMargin: -50.2,
  totalCostPerAttendee: 1047,
  netCostPerAttendee: 373
};

const DISCOUNT_CAPS = {
  'REVGUEST-A': 20,
  'REVGUEST-B': 20,
  REVFAN100: 15,
  'REVSALES50-EB': 20,
  SUMMIT100VIIRTUE: null,
  CAMREVSUMMIT: null
};

function loadEnv() {
  const envPath = path.join(workspaceDir, '.env');
  if (!fs.existsSync(envPath)) return;

  const env = fs.readFileSync(envPath, 'utf8');
  for (const line of env.split('\n')) {
    const [key, ...value] = line.split('=');
    if (key && value.length && !process.env[key.trim()]) {
      process.env[key.trim()] = value.join('=').trim();
    }
  }
}

function sfRequest({ method = 'GET', requestPath, token, body, headers = {} }) {
  const instanceUrl = process.env.SF_INSTANCE_URL;
  if (!instanceUrl) throw new Error('SF_INSTANCE_URL is not configured.');

  return new Promise((resolve, reject) => {
    const request = https.request({
      hostname: new URL(instanceUrl).hostname,
      path: requestPath,
      method,
      headers: {
        ...(token ? { Authorization: `Bearer ${token}` } : {}),
        ...headers
      }
    }, (response) => {
      let data = '';
      response.on('data', (chunk) => data += chunk);
      response.on('end', () => {
        let parsed = data;
        try {
          parsed = data ? JSON.parse(data) : null;
        } catch {
          // Preserve raw text for Salesforce diagnostic errors.
        }

        if (response.statusCode >= 400) {
          const message = Array.isArray(parsed)
            ? parsed.map((item) => item.message || item.errorCode).join('; ')
            : parsed?.message || parsed?.error_description || data || `HTTP ${response.statusCode}`;
          reject(new Error(message));
          return;
        }

        resolve(parsed);
      });
    });

    request.on('error', reject);
    if (body) request.write(body);
    request.end();
  });
}

async function getSalesforceToken() {
  loadEnv();
  if (!process.env.SF_CLIENT_ID || !process.env.SF_CLIENT_SECRET || !process.env.SF_INSTANCE_URL) {
    console.warn('Salesforce credentials not found; building without sponsorship opportunities.');
    return null;
  }

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: process.env.SF_CLIENT_ID,
    client_secret: process.env.SF_CLIENT_SECRET
  }).toString();

  const parsed = await sfRequest({
    method: 'POST',
    requestPath: '/services/oauth2/token',
    body,
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(body)
    }
  });

  return parsed.access_token;
}

async function sfQueryAll(token, soql) {
  let result = await sfRequest({
    token,
    requestPath: `/services/data/${API_VERSION}/query?q=${encodeURIComponent(soql)}`
  });
  let records = result.records || [];

  while (result.nextRecordsUrl) {
    result = await sfRequest({ token, requestPath: result.nextRecordsUrl });
    records = records.concat(result.records || []);
  }

  return records;
}

function soqlString(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function valueOf(cell) {
  if (cell == null) return '';
  if (cell instanceof Date) return cell.toISOString();
  if (typeof cell === 'object') return cell.text || cell.result || cell.hyperlink || '';
  return cell;
}

function asText(value) {
  return String(valueOf(value) ?? '')
    .replace(/\u0096/g, '–')
    .replace(/\u0092/g, "'")
    .trim();
}

function asNumber(value) {
  if (typeof value === 'number') return value;
  const text = asText(value).replace(/[$,]/g, '');
  const num = Number(text);
  return Number.isFinite(num) ? num : 0;
}

function parseDate(value) {
  if (value instanceof Date) return value;
  const text = asText(value);
  if (!text) return null;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function cleanPackage(value) {
  return asText(value)
    .replace(/^Standard:\s*/i, '')
    .replace(/\s*\([^)]*\)/g, '')
    .replace(/-\s*$/g, '')
    .replace(/\s+/g, ' ')
    .trim() || 'Unspecified';
}

function cleanIndustry(value) {
  return asText(value)
    .replace('Telecom or Communications Service Provider (CSP)', 'CSP / Telecom')
    .replace('Managed Service Provider (MSP)', 'MSP')
    .trim() || 'Unspecified';
}

function pct(current, goal) {
  return goal ? Math.round((current / goal) * 100) : 0;
}

function money(value) {
  const abs = Math.abs(value);
  const formatted = '$' + Math.round(abs).toLocaleString('en-US');
  return value < 0 ? `(${formatted})` : formatted;
}

function countBy(rows, selector) {
  const counts = new Map();
  rows.forEach((row) => {
    const raw = selector(row);
    if (!raw) return;
    const values = Array.isArray(raw) ? raw : [raw];
    values.forEach((value) => {
      const key = String(value).trim();
      if (!key) return;
      counts.set(key, (counts.get(key) || 0) + 1);
    });
  });
  return [...counts.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || a.name.localeCompare(b.name));
}

function splitMulti(value) {
  return asText(value).split(',').map((item) => item.trim()).filter(Boolean);
}

function normalizeTopics(value) {
  const text = asText(value);
  const topics = [
    'Business Growth, Strategy & Service Delivery',
    'Customer Success & Experience',
    'Industry & Regulatory Insights',
    'Sales & Marketing',
    'Technology, AI & Innovation',
    'Product & Platform Deep Dives',
    'Communications Billing (Usage rating, Telecom Taxation)',
    'Payment Processing (Accounts receivable, Accounts Payable)',
    'PSA (service tickets, tech scheduling, inventory, quoting)',
    'RMM (Remote Monitoring and Management)',
    'Women in Leadership'
  ];
  return topics.filter((topic) => text.includes(topic));
}

function normalizeReceptions(value) {
  const text = asText(value);
  const receptions = [];
  if (text.includes('Welcome Reception')) receptions.push('Welcome Reception');
  if (text.includes('Networking Reception')) receptions.push('Networking Reception');
  if (text.includes('not attending')) receptions.push('Not Attending Receptions');
  return receptions.length ? receptions : ['Unspecified'];
}

function cleanReferral(firstName, lastName) {
  const parts = [firstName, lastName]
    .map(asText)
    .filter((part) => part && !/^n\/?a$/i.test(part));
  return parts.join(' ');
}

function normalizeRegistrationRecord(record) {
  const firstName = asText(record['First Name (1557966)']);
  const lastName = asText(record['Last Name (1557967)']);
  const email = asText(record.Email);
  if (!firstName && !lastName && !email) return null;

  const date = parseDate(record['Date Registered']);
  const discountCode = asText(record['Discount Code']).toUpperCase();
  const packageAmount = asNumber(record['Package Amount']);
  const discountAmount = asNumber(record['Discount Amount Val']);
  const amount = asNumber(record.Amount);
  const userType = asText(record['User Type']);

  return {
    name: `${firstName} ${lastName}`.trim() || email,
    company: asText(record.Company) || 'Unspecified',
    jobTitle: asText(record['Job Title']),
    state: asText(record['State/ Province']),
    dateRegistered: date ? date.toISOString().slice(0, 10) : '',
    dateLabel: date ? date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric', timeZone: 'UTC' }) : '',
    status: asText(record.Status),
    hasPaid: /^yes$/i.test(asText(record['Has Paid'])),
    userType,
    isStaff: /@rev\.io$/i.test(email) || /staff/i.test(userType),
    isSponsorPass: /sponsor/i.test(userType) || /sponsor/i.test(asText(record['Package Name'])),
    businessType: cleanIndustry(record['Which Of The Following Best Describes Your Business?']),
    jobRank: asText(record['Job Rank']) || 'Unspecified',
    department: asText(record.Department) || 'Unspecified',
    companySize: asText(record['Company Size']) || 'Unspecified',
    packageName: cleanPackage(record['Package Name']),
    packageAmount,
    discountCode,
    discountAmount,
    amount,
    paymentStatus: asText(record['Payment Status']) || 'Unspecified',
    referral: cleanReferral(record['Referral First Name'], record['Referral Last Name']),
    receptions: normalizeReceptions(record['Which Receptions Are You Planning To Attend?']),
    hotel: asText(record['Do You Need A Hotel Room For The 2026 Rev.io Summit?']) || 'Unspecified',
    shirtSize: asText(record['T Shirt Size']) || 'Unspecified',
    attendedBefore: asText(record['Have You Attended The Rev.io Summit Before?']) || 'Unspecified',
    topics: normalizeTopics(record['Please Select The Topics That Best Match Your Areas Of Interest. (Select All That Apply.)'])
  };
}

async function loadExcelRecords() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(workbookPath);
  const sheet = workbook.worksheets[0];
  const headers = sheet.getRow(1).values.slice(1).map(asText);
  const records = [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const record = {};
    headers.forEach((header, index) => {
      record[header] = valueOf(row.getCell(index + 1).value);
    });
    records.push(record);
  });

  return records;
}

function loadCsvRecords() {
  const workbook = XLSX.readFile(workbookPath, { codepage: 1252, raw: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

async function loadWorkbookRows() {
  const records = path.extname(workbookPath).toLowerCase() === '.csv'
    ? loadCsvRecords()
    : await loadExcelRecords();

  return records
    .map(normalizeRegistrationRecord)
    .filter(Boolean)
    .filter((row) => !/cancel/i.test(row.status));
}

async function loadSponsorshipOpportunities() {
  const token = await getSalesforceToken();
  if (!token) return [];

  const records = await sfQueryAll(token, `
    SELECT Id, Name, Account.Name, StageName, IsClosed, IsWon, CloseDate,
           Sponsorship_Amount__c, Sponsorship_Type__c, Sponsorship_Package__c
    FROM Opportunity
    WHERE Type = '${soqlString(SPONSOR_OPPORTUNITY_TYPE)}'
    ORDER BY CloseDate DESC, Account.Name ASC
  `);

  return records.map((record) => ({
    opportunityName: record.Name || '',
    company: record.Account?.Name || 'Unspecified',
    stage: record.StageName || 'Unspecified',
    closeDate: record.CloseDate || '',
    amount: Number(record.Sponsorship_Amount__c || 0),
    sponsorshipType: record.Sponsorship_Type__c || 'Unspecified',
    sponsorshipPackage: record.Sponsorship_Package__c || '',
    isClosed: Boolean(record.IsClosed),
    isWon: Boolean(record.IsWon)
  }));
}

function buildSponsorshipSummary(opportunities) {
  const uniqueCompanies = new Set(opportunities.map((row) => row.company.toLowerCase()).filter((value) => value && value !== 'unspecified')).size;
  const totalAmount = opportunities.reduce((sum, row) => sum + row.amount, 0);
  const wonAmount = opportunities.filter((row) => row.isWon).reduce((sum, row) => sum + row.amount, 0);
  const openAmount = opportunities.filter((row) => !row.isClosed).reduce((sum, row) => sum + row.amount, 0);

  return {
    sourceOpportunityType: SPONSOR_OPPORTUNITY_TYPE,
    totalOpportunities: opportunities.length,
    uniqueCompanies,
    totalAmount,
    wonAmount,
    openAmount,
    closedWon: opportunities.filter((row) => row.isWon).length,
    closedLost: opportunities.filter((row) => row.isClosed && !row.isWon).length,
    byStage: countBy(opportunities, (row) => row.stage),
    byType: countBy(opportunities, (row) => row.sponsorshipType === 'Unspecified' ? null : row.sponsorshipType)
  };
}

function buildModel(rows, sponsorshipOpportunities) {
  const uniqueCompanies = new Set(rows.map((row) => row.company.toLowerCase()).filter((v) => v && v !== 'unspecified')).size;
  const staffRows = rows.filter((row) => row.isStaff);
  const sponsorRows = rows.filter((row) => row.isSponsorPass);
  const customerRows = rows.filter((row) => !row.isStaff && !row.isSponsorPass);
  const paidRows = customerRows.filter((row) => row.amount > 0);
  const compRows = customerRows.filter((row) => row.amount === 0);
  const ticketRevenue = rows.reduce((sum, row) => sum + row.amount, 0);
  const sponsorshipSummary = buildSponsorshipSummary(sponsorshipOpportunities);
  const avgPaidTicket = paidRows.length ? Math.round(ticketRevenue / paidRows.length) : 0;
  const fullComps = rows.filter((row) => row.amount === 0 && (row.discountAmount >= row.packageAmount || row.discountCode));
  const referrals = rows.filter((row) => row.referral).map((row, index) => ({
    id: index + 1,
    registrant: row.name,
    referrer: row.referral,
    date: row.dateRegistered
  }));

  const ticketTypes = countBy(rows, (row) => row.packageName).map((entry) => {
    const matching = rows.filter((row) => row.packageName === entry.name);
    return {
      ...entry,
      paid: matching.filter((row) => row.amount > 0).length,
      comp: matching.filter((row) => row.amount === 0).length,
      revenue: matching.reduce((sum, row) => sum + row.amount, 0)
    };
  });

  const discounts = countBy(rows, (row) => row.discountCode || null).map((entry) => {
    const cap = Object.prototype.hasOwnProperty.call(DISCOUNT_CAPS, entry.name) ? DISCOUNT_CAPS[entry.name] : null;
    return { ...entry, cap, remaining: cap == null ? null : Math.max(0, cap - entry.count) };
  });

  const byMonth = countBy(rows, (row) => {
    if (!row.dateRegistered) return null;
    const date = new Date(row.dateRegistered + 'T00:00:00Z');
    return date.toLocaleDateString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
  });

  return {
    generatedAt: new Date().toISOString(),
    sourceFile: path.basename(workbookPath),
    goals: GOALS,
    actual2025: ACTUAL_2025,
    summary: {
      totalRegistrants: rows.length,
      uniqueCompanies,
      customerProspectRegistrations: customerRows.length,
      paidTickets: paidRows.length,
      compTickets: compRows.length,
      staffTickets: staffRows.length,
      sponsorPasses: sponsorRows.length,
      ticketRevenue,
      sponsorRevenue: sponsorshipSummary.totalAmount,
      totalRevenue: ticketRevenue + sponsorshipSummary.totalAmount,
      avgPaidTicket,
      fullCompTickets: fullComps.length,
      hotelRooms: rows.filter((row) => /^yes$/i.test(row.hotel)).length,
      attendedBefore: rows.filter((row) => /^yes$/i.test(row.attendedBefore)).length,
      referralCount: referrals.length
    },
    registrants: rows.map((row) => ({
      name: row.name,
      company: row.company,
      jobTitle: row.jobTitle,
      state: row.state,
      dateLabel: row.dateLabel,
      dateRegistered: row.dateRegistered,
      businessType: row.businessType,
      jobRank: row.jobRank,
      department: row.department,
      companySize: row.companySize,
      packageName: row.packageName,
      amount: row.amount,
      discountCode: row.discountCode,
      referral: row.referral,
      hotel: row.hotel,
      shirtSize: row.shirtSize,
      attendedBefore: row.attendedBefore
    })),
    referrals,
    agenda: AGENDA_DAYS,
    sponsorships: {
      ...sponsorshipSummary,
      opportunities: sponsorshipOpportunities
    },
    breakdowns: {
      ticketTypes,
      discounts,
      businessTypes: countBy(rows, (row) => row.businessType),
      jobRanks: countBy(rows, (row) => row.jobRank),
      departments: countBy(rows, (row) => row.department),
      companySizes: countBy(rows, (row) => row.companySize),
      states: countBy(rows, (row) => row.state || null),
      hotel: countBy(rows, (row) => row.hotel),
      shirts: countBy(rows, (row) => row.shirtSize),
      topics: countBy(rows, (row) => row.topics),
      registrationsByMonth: byMonth,
      receptionPlans: countBy(rows, (row) => row.receptions)
    }
  };
}

function renderDashboard(model) {
  const json = JSON.stringify(model).replace(/</g, '\\u003c');
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>2026 Rev.io Client Summit Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
:root {
  --ink-900: #0a141f;
  --ink-850: #0c1925;
  --ink-800: #0f1b2a;
  --ink-700: #15283f;
  --navy: #1D3756;
  --navy-dark: #10243a;
  --navy-deep: #071726;
  --teal: #2399B5;
  --cyan: #34bde5;
  --cyan-soft: #7fd9ef;
  --aurora-teal: #4fd1c5;
  --aurora-lime: #c6f178;
  --green: #6EBE4F;
  --card: rgba(15, 27, 42, 0.88);
  --card-strong: rgba(21, 40, 63, 0.94);
  --line: rgba(255, 255, 255, 0.11);
  --line-strong: rgba(127, 217, 239, 0.24);
  --text: #FFFFFF;
  --muted: #b9c7d6;
  --dim: #8ea3b9;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  min-height: 100vh;
  background: var(--ink-900);
  color: var(--text);
  font-family: "Open Sans", Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
}
body::before {
  content: "";
  position: fixed;
  inset: 0;
  pointer-events: none;
  background:
    linear-gradient(115deg, rgba(52,189,229,.16) 0%, transparent 34%),
    linear-gradient(250deg, rgba(198,241,120,.12) 0%, transparent 38%),
    linear-gradient(180deg, var(--ink-900) 0%, var(--ink-850) 42%, #060e18 100%);
}
body.auth-locked { overflow: hidden; }
body.auth-locked .app { display: none; }
.app { position: relative; z-index: 1; }
.auth-gate {
  position: fixed;
  inset: 0;
  z-index: 100;
  display: grid;
  place-items: center;
  padding: 24px;
  background:
    linear-gradient(180deg, rgba(10,20,31,.18), rgba(10,20,31,.94)),
    radial-gradient(ellipse at 65% 18%, rgba(52,189,229,.22), transparent 34%),
    var(--ink-900);
}
.auth-gate.hidden { display: none; }
.auth-panel {
  width: min(420px, 100%);
  border: 1px solid var(--line-strong);
  border-radius: 8px;
  background:
    radial-gradient(circle at 88% 8%, rgba(52,189,229,.18), transparent 34%),
    linear-gradient(145deg, rgba(21,40,63,.96), rgba(10,20,31,.9));
  box-shadow: 0 24px 60px rgba(0,0,0,.38);
  padding: 26px;
}
.auth-panel h1 {
  margin: 14px 0 8px;
  color: white;
  font-size: 30px;
  line-height: 1.05;
  text-transform: uppercase;
}
.auth-panel p {
  margin: 0 0 18px;
  color: var(--muted);
  line-height: 1.45;
}
.auth-form {
  display: grid;
  gap: 12px;
}
.auth-form label {
  color: var(--cyan-soft);
  font-size: 11px;
  font-weight: 800;
  letter-spacing: 1.4px;
  text-transform: uppercase;
}
.auth-form input {
  min-height: 46px;
  border: 1px solid rgba(255,255,255,.16);
  border-radius: 8px;
  background: rgba(10,20,31,.76);
  color: white;
  padding: 0 13px;
  font: 700 16px "Open Sans", Arial, sans-serif;
}
.auth-form input:focus {
  border-color: var(--cyan);
  outline: 3px solid rgba(52,189,229,.18);
}
.auth-form button {
  min-height: 44px;
  border: 0;
  border-radius: 8px;
  background: linear-gradient(135deg, var(--aurora-lime), var(--cyan));
  color: white;
  padding: 0 18px;
  font: 800 13px "Open Sans", Arial, sans-serif;
  cursor: pointer;
}
.auth-error {
  min-height: 18px;
  color: white;
  font-size: 12px;
  font-weight: 800;
}
.hero {
  position: relative;
  min-height: 360px;
  display: grid;
  align-items: end;
  overflow: hidden;
  background:
    linear-gradient(180deg, rgba(10,20,31,.2), rgba(10,20,31,.96)),
    radial-gradient(ellipse at 66% 16%, rgba(52,189,229,.22), transparent 34%);
}
.hero-art {
  position: absolute;
  inset: 0;
  z-index: -1;
  opacity: .98;
}
.hero-inner {
  width: min(1400px, calc(100% - 40px));
  margin: 0 auto;
  padding: 82px 0 36px;
}
.eyebrow {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  border: 1px solid rgba(127,217,239,.42);
  color: white;
  background: rgba(52,189,229,.12);
  padding: 7px 13px;
  border-radius: 999px;
  font-size: 11px;
  letter-spacing: 2.2px;
  text-transform: uppercase;
  font-weight: 700;
}
.eyebrow::before {
  content: "";
  width: 7px;
  height: 7px;
  border-radius: 999px;
  background: var(--cyan);
  box-shadow: 0 0 0 4px rgba(52,189,229,.18), 0 0 16px rgba(52,189,229,.75);
}
h1, h2, h3 { font-family: Montserrat, "Open Sans", sans-serif; }
.hero h1 {
  max-width: 920px;
  margin: 18px 0 12px;
  font-size: clamp(44px, 7vw, 88px);
  line-height: .96;
  letter-spacing: 0;
  text-transform: uppercase;
}
.hero-subtitle {
  margin: 0 0 16px;
  color: white;
  font-size: clamp(18px, 2.4vw, 30px);
  font-weight: 700;
}
.hero-meta {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  color: white;
  font-weight: 700;
  letter-spacing: 2px;
  text-transform: uppercase;
}
.hero-meta span {
  border: 1px solid rgba(255,255,255,.14);
  border-radius: 999px;
  background: rgba(10,20,31,.42);
  padding: 8px 12px;
  font-size: 12px;
}
.tabs {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  justify-content: center;
  gap: 0;
  background: rgba(10,20,31,.9);
  border-bottom: 1px solid var(--line-strong);
  backdrop-filter: blur(14px);
  scrollbar-width: none;
}
.tabs::-webkit-scrollbar { display: none; }
.tab {
  appearance: none;
  border: 0;
  border-bottom: 3px solid transparent;
  background: transparent;
  color: var(--muted);
  padding: 15px 28px 13px;
  font: 800 12px "Open Sans", Arial, sans-serif;
  letter-spacing: 1px;
  text-transform: uppercase;
  cursor: pointer;
}
.tab:hover { color: white; background: rgba(127,217,239,.08); }
.tab.active { color: white; border-bottom-color: var(--cyan); background: rgba(52,189,229,.12); }
main { width: min(1400px, calc(100% - 40px)); margin: 0 auto; padding: 34px 0 48px; }
.panel { display: none; }
.panel.active { display: block; }
.section-title {
  display: flex;
  align-items: center;
  gap: 12px;
  margin: 34px 0 16px;
  color: white;
  font-size: 13px;
  letter-spacing: 2.2px;
  text-transform: uppercase;
}
.section-title:first-child { margin-top: 0; }
.section-title::before { content: attr(data-num); color: var(--cyan); font-family: Montserrat, "Open Sans", sans-serif; font-size: 12px; font-weight: 800; }
.section-title::after { content: ""; flex: 1; max-width: 120px; height: 1px; background: var(--line-strong); }
.metric-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 16px; }
.metric-card, .chart-card {
  border: 1px solid var(--line);
  border-radius: 8px;
  background:
    linear-gradient(145deg, rgba(21,40,63,.94), rgba(10,20,31,.82)),
    linear-gradient(90deg, rgba(52,189,229,.08), rgba(110,190,79,.06));
  box-shadow: 0 18px 36px rgba(0,0,0,.24);
}
.metric-card { padding: 20px; min-height: 138px; }
.metric-label { color: var(--cyan-soft); font-size: 11px; font-weight: 800; letter-spacing: 1.4px; text-transform: uppercase; }
.metric-value { margin-top: 9px; font-family: Montserrat, "Open Sans", sans-serif; font-size: 36px; font-weight: 800; line-height: 1; overflow-wrap: anywhere; }
.metric-note { margin-top: 10px; color: var(--muted); font-size: 12px; }
.bar-track { height: 8px; margin-top: 16px; overflow: hidden; border-radius: 999px; background: rgba(255,255,255,.12); }
.bar-fill { height: 100%; border-radius: inherit; background: linear-gradient(90deg, var(--aurora-lime), var(--cyan)); }
.chart-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
.chart-card { padding: 18px; }
.chart-title { margin: 0 0 14px; font-size: 14px; text-transform: uppercase; letter-spacing: .8px; }
.row-bar { display: grid; grid-template-columns: minmax(130px, 1fr) minmax(110px, 2fr) 42px; gap: 10px; align-items: center; margin: 10px 0; font-size: 13px; }
.row-name { color: white; overflow-wrap: anywhere; }
.mini-track { height: 8px; border-radius: 999px; background: rgba(255,255,255,.1); overflow: hidden; }
.mini-fill { height: 100%; border-radius: inherit; background: linear-gradient(90deg, var(--cyan), var(--aurora-teal)); }
.row-count { color: white; text-align: right; font-weight: 800; }
.click-row {
  width: 100%;
  border: 0;
  border-radius: 8px;
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: left;
  cursor: pointer;
}
.click-row:hover, .click-row:focus-visible {
  background: rgba(127,217,239,.08);
  outline: 1px solid rgba(127,217,239,.22);
}
.pill.click-row { width: auto; }
.table-wrap {
  overflow: auto;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(10,20,31,.72);
}
table { width: 100%; border-collapse: collapse; min-width: 820px; }
th, td { padding: 11px 13px; border-bottom: 1px solid rgba(255,255,255,.09); text-align: left; vertical-align: top; }
th {
  position: sticky;
  top: 0;
  color: white;
  background: #15283f;
  font-size: 11px;
  letter-spacing: .7px;
  text-transform: uppercase;
  z-index: 1;
}
td { color: white; font-size: 13px; }
.num { text-align: right; font-variant-numeric: tabular-nums; }
.muted { color: var(--muted); }
.pill { display: inline-flex; padding: 3px 9px; border-radius: 999px; border: 1px solid rgba(127,217,239,.28); background: rgba(52,189,229,.12); font-size: 11px; font-weight: 800; color: white; }
.sort-button {
  display: inline-flex;
  align-items: center;
  gap: 6px;
  border: 0;
  background: transparent;
  color: white;
  padding: 0;
  font: inherit;
  font-weight: 800;
  letter-spacing: inherit;
  text-transform: inherit;
  cursor: pointer;
}
.sort-button::after { content: '↕'; color: rgba(255,255,255,.46); font-size: 10px; }
.sort-button.active.asc::after { content: '↑'; color: var(--cyan-soft); }
.sort-button.active.desc::after { content: '↓'; color: var(--cyan-soft); }
.filter-row th { top: 37px; background: #0f1d2d; padding: 7px 8px; z-index: 2; }
.column-filter {
  width: 100%;
  min-height: 32px;
  border: 1px solid rgba(255,255,255,.14);
  border-radius: 7px;
  background: rgba(10,20,31,.88);
  color: white;
  padding: 0 8px;
  font: 12px "Open Sans", sans-serif;
  text-transform: none;
  letter-spacing: 0;
}
.column-filter::placeholder { color: rgba(255,255,255,.42); }
.toolbar { display: flex; gap: 10px; margin-bottom: 14px; flex-wrap: wrap; }
.toolbar input, .toolbar select {
  min-height: 40px;
  border: 1px solid rgba(255,255,255,.16);
  border-radius: 8px;
  background: rgba(10,20,31,.76);
  color: white;
  padding: 0 12px;
  font: inherit;
}
.toolbar input { flex: 1; min-width: 260px; }
.active-filter {
  display: none;
  align-items: center;
  gap: 8px;
  width: 100%;
  color: white;
  font-size: 12px;
  font-weight: 800;
}
.active-filter.visible { display: flex; }
.action-button {
  min-height: 40px;
  border: 0;
  border-radius: 8px;
  background: linear-gradient(135deg, var(--aurora-lime), var(--cyan));
  color: white;
  padding: 0 18px;
  font: 800 13px "Open Sans", Arial, sans-serif;
  cursor: pointer;
}
.subtabs { display: flex; flex-wrap: wrap; border-bottom: 1px solid var(--line); margin-bottom: 18px; }
.subtab {
  border: 0;
  border-bottom: 3px solid transparent;
  background: transparent;
  color: var(--muted);
  padding: 12px 18px 10px;
  font-weight: 800;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: .8px;
  cursor: pointer;
}
.subtab:hover { color: white; background: rgba(127,217,239,.08); }
.subtab.active { color: white; border-bottom-color: var(--cyan); }
.subpanel { display: none; }
.subpanel.active { display: block; }
.agenda-hero {
  display: grid;
  grid-template-columns: minmax(0, 1.4fr) minmax(280px, .6fr);
  gap: 18px;
  align-items: stretch;
}
.agenda-intro {
  padding: 22px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background:
    radial-gradient(circle at 86% 8%, rgba(52,189,229,.18), transparent 32%),
    linear-gradient(145deg, rgba(21,40,63,.94), rgba(10,20,31,.82));
}
.agenda-intro h3 {
  margin: 10px 0 10px;
  font-size: clamp(26px, 4vw, 46px);
  line-height: 1;
  text-transform: uppercase;
}
.agenda-intro p { max-width: 760px; margin: 0; color: white; font-size: 15px; line-height: 1.55; }
.agenda-actions {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
  margin-top: 18px;
}
.agenda-download {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  border-radius: 8px;
  background: linear-gradient(135deg, var(--aurora-lime), var(--cyan));
  color: white;
  padding: 0 18px;
  font: 800 13px "Open Sans", Arial, sans-serif;
  text-decoration: none;
}
.agenda-stats { display: grid; grid-template-columns: 1fr; gap: 12px; }
.agenda-day {
  margin-top: 24px;
  border: 1px solid var(--line);
  border-radius: 8px;
  overflow: hidden;
  background: rgba(10,20,31,.62);
}
.agenda-day-header {
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: 18px;
  padding: 18px 20px;
  background: linear-gradient(90deg, rgba(29,55,86,.98), rgba(35,153,181,.26));
  border-bottom: 1px solid var(--line-strong);
}
.agenda-day-label {
  color: var(--cyan-soft);
  font-size: 11px;
  font-weight: 800;
  letter-spacing: 1.8px;
  text-transform: uppercase;
}
.agenda-day-title { margin-top: 4px; color: white; font-family: Montserrat, "Open Sans", sans-serif; font-size: 21px; font-weight: 800; }
.agenda-day-theme { color: white; font-weight: 800; text-align: right; }
.agenda-list { display: grid; gap: 0; }
.agenda-item {
  display: grid;
  grid-template-columns: 160px minmax(0, 1fr) 210px 120px;
  gap: 14px;
  padding: 15px 18px;
  border-bottom: 1px solid rgba(255,255,255,.08);
}
.agenda-item:last-child { border-bottom: 0; }
.agenda-time { color: white; font-weight: 800; font-variant-numeric: tabular-nums; }
.agenda-title { color: white; font-weight: 800; line-height: 1.35; overflow-wrap: anywhere; }
.agenda-meta { color: var(--muted); font-size: 12px; line-height: 1.45; }
.agenda-room { color: white; font-size: 12px; font-weight: 800; text-align: right; }
.ops-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
.ops-card {
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(10,20,31,.72);
  padding: 18px;
}
.ops-card h3 {
  margin: 0 0 12px;
  color: white;
  font-size: 16px;
  font-weight: 800;
}
.ops-list { display: grid; gap: 10px; margin: 0; padding: 0; list-style: none; }
.ops-list li {
  display: grid;
  grid-template-columns: 122px minmax(0, 1fr);
  gap: 12px;
  align-items: start;
  padding-bottom: 10px;
  border-bottom: 1px solid rgba(255,255,255,.08);
  color: white;
  line-height: 1.42;
}
.ops-list li:last-child { border-bottom: 0; padding-bottom: 0; }
.ops-time {
  color: var(--cyan-soft);
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 12px;
  font-weight: 800;
  font-variant-numeric: tabular-nums;
}
.ops-detail strong { display: block; color: white; font-weight: 800; }
.ops-detail span { display: block; margin-top: 3px; color: var(--muted); font-size: 12px; }
.link-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
.link-card {
  display: block;
  min-height: 132px;
  border: 1px solid var(--line);
  border-radius: 8px;
  background: rgba(10,20,31,.72);
  padding: 18px;
  color: white;
  text-decoration: none;
}
.link-card:hover { border-color: rgba(127,217,239,.6); background: rgba(18,39,61,.82); }
.link-card h3 { margin: 0 0 8px; color: white; font-size: 17px; }
.link-card p { margin: 0; color: var(--muted); line-height: 1.45; }
.link-url {
  display: block;
  margin-top: 14px;
  color: var(--cyan-soft);
  font-size: 12px;
  font-weight: 800;
  overflow-wrap: anywhere;
}
.footer { border-top: 1px solid var(--line); color: var(--muted); font-size: 12px; text-align: center; padding: 26px 20px 36px; position: relative; z-index: 1; }
@media (max-width: 900px) {
  .metric-grid, .chart-grid, .ops-grid, .link-grid { grid-template-columns: 1fr; }
  .agenda-hero { grid-template-columns: 1fr; }
  .agenda-item { grid-template-columns: 1fr; gap: 5px; }
  .ops-list li { grid-template-columns: 1fr; gap: 4px; }
  .agenda-room, .agenda-day-theme { text-align: left; }
  .tabs { overflow-x: auto; justify-content: flex-start; }
  .tab { white-space: nowrap; }
  main, .hero-inner { width: min(100% - 28px, 1400px); }
  .hero { min-height: 330px; }
  .hero-meta span { font-size: 11px; }
}
</style>
</head>
<body class="auth-locked">
<div class="auth-gate" id="authGate">
  <section class="auth-panel" aria-labelledby="authTitle">
    <div class="eyebrow">Internal Access</div>
    <h1 id="authTitle">Rev.io Client Summit</h1>
    <p>Enter the dashboard password to continue.</p>
    <form class="auth-form" id="authForm">
      <label for="dashboardPassword">Password</label>
      <input id="dashboardPassword" type="password" autocomplete="current-password" autofocus>
      <button type="submit">Open Dashboard</button>
      <div class="auth-error" id="authError" aria-live="polite"></div>
    </form>
  </section>
</div>
<div class="app" id="app">
  <header class="hero">
    <svg class="hero-art" viewBox="0 0 1400 420" preserveAspectRatio="none" aria-hidden="true">
      <defs>
        <linearGradient id="summitSky" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#060e18"/><stop offset=".52" stop-color="#15283f"/><stop offset="1" stop-color="#2399B5"/></linearGradient>
        <linearGradient id="aurora" x1="0" y1="0" x2="1" y2="0"><stop stop-color="#34bde5" stop-opacity=".03"/><stop offset=".45" stop-color="#4fd1c5" stop-opacity=".38"/><stop offset="1" stop-color="#c6f178" stop-opacity=".14"/></linearGradient>
        <linearGradient id="farRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#294b72"/><stop offset="1" stop-color="#0f1b2a"/></linearGradient>
        <linearGradient id="nearRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#1d3756"/><stop offset="1" stop-color="#060e18"/></linearGradient>
      </defs>
      <rect width="1400" height="420" fill="url(#summitSky)"/>
      <path d="M-40 84 C180 8 350 92 508 42 C674 -10 795 52 940 24 C1112 -9 1262 38 1440 4 L1440 176 C1262 198 1116 158 960 184 C780 216 662 146 502 196 C328 250 158 174 -40 244Z" fill="url(#aurora)" opacity=".9"/>
      <g opacity=".58">
        <circle cx="176" cy="74" r="1.6" fill="#f5f9ff"/><circle cx="328" cy="132" r="1.2" fill="#f5f9ff"/><circle cx="512" cy="76" r="1.4" fill="#f5f9ff"/><circle cx="746" cy="104" r="1.1" fill="#f5f9ff"/><circle cx="1018" cy="64" r="1.5" fill="#f5f9ff"/><circle cx="1212" cy="124" r="1.1" fill="#f5f9ff"/>
      </g>
      <path d="M0 296 L86 244 L172 270 L276 198 L372 246 L484 176 L596 250 L702 150 L816 238 L916 182 L1030 250 L1138 162 L1246 232 L1328 196 L1400 224 L1400 420 L0 420Z" fill="url(#farRidge)" opacity=".86"/>
      <path d="M0 344 L126 298 L236 318 L354 254 L456 306 L570 232 L700 318 L842 218 L982 314 L1112 244 L1232 304 L1330 270 L1400 306 L1400 420 L0 420Z" fill="url(#nearRidge)" opacity=".96"/>
      <path d="M0 344 L126 298 L236 318 L354 254 L456 306 L570 232 L700 318 L842 218 L982 314 L1112 244 L1232 304 L1330 270 L1400 306" fill="none" stroke="#7fd9ef" stroke-width="1.2" opacity=".34"/>
      <path d="M0 382 C178 354 322 392 496 360 C660 330 808 384 960 352 C1146 312 1268 354 1400 328 L1400 420 L0 420Z" fill="#060e18" opacity=".82"/>
    </svg>
    <div class="hero-inner">
      <div class="eyebrow">Internal Team Dashboard</div>
      <h1>Rev.io Client Summit</h1>
      <p class="hero-subtitle">Prepare for Tomorrow.</p>
      <div class="hero-meta"><span>September 1-3, 2026</span><span>Atlanta, GA</span><span>Live registration + Salesforce sponsorship view</span></div>
    </div>
  </header>
  <nav class="tabs" aria-label="Dashboard sections">
    <button class="tab active" data-tab="home">Home</button>
    <button class="tab" data-tab="registrants">Registrants</button>
    <button class="tab" data-tab="agenda">Agenda</button>
    <button class="tab" data-tab="onsite">Onsite Ops</button>
    <button class="tab" data-tab="links">Important Links</button>
    <button class="tab" data-tab="referrals">Referrals</button>
    <button class="tab" data-tab="sponsorships">Sponsorships</button>
  </nav>
  <main>
    <section class="panel active" id="home">
      <h2 class="section-title" data-num="01">Registration Goals</h2>
      <div class="metric-grid" id="registrationCards"></div>
      <h2 class="section-title" data-num="02">Revenue Goals</h2>
      <div class="metric-grid" id="revenueCards"></div>
      <h2 class="section-title" data-num="03">2025 Actuals vs 2026 Current</h2>
      <div class="table-wrap"><table id="modelTable"></table></div>
      <h2 class="section-title" data-num="04">Audience Mix</h2>
      <div class="chart-grid">
        <div class="chart-card"><h3 class="chart-title">Business Type</h3><div id="businessTypes"></div></div>
        <div class="chart-card"><h3 class="chart-title">Department</h3><div id="departments"></div></div>
        <div class="chart-card"><h3 class="chart-title">Job Rank</h3><div id="jobRanks"></div></div>
        <div class="chart-card"><h3 class="chart-title">Registration Timeline</h3><div id="registrationsByMonth"></div></div>
      </div>
    </section>
    <section class="panel" id="registrants">
      <h2 class="section-title" data-num="01">Registration Summary</h2>
      <div class="metric-grid" id="registrantCards"></div>
      <div class="toolbar">
        <input id="search" type="search" placeholder="Search name, company, title, discount code">
        <select id="ticketFilter"><option value="">All ticket types</option></select>
        <select id="companyFilter"><option value="">All companies</option></select>
        <select id="referrerFilter"><option value="">All referrers</option></select>
        <div id="activeRegistrantFilter" class="active-filter"></div>
      </div>
      <div class="table-wrap"><table id="registrantsTable"></table></div>
      <h2 class="section-title" data-num="02">Ticket Pricing</h2>
      <div class="table-wrap"><table id="ticketTable"></table></div>
    </section>
    <section class="panel" id="agenda">
      <div class="agenda-hero">
        <div class="agenda-intro">
          <div class="eyebrow">Team-Shareable Agenda</div>
          <h3>Summit Agenda</h3>
          <p>Three days built around the next operating model for service providers: market perspective, Rev.io platform strategy, AI-enabled workflows, customer operations, and sponsor-led learning moments.</p>
          <div class="agenda-actions">
            <a class="agenda-download" href="assets/Revio-Summit-2026-Agenda.pdf" target="_blank" rel="noopener">Open Shareable PDF</a>
          </div>
        </div>
        <div class="agenda-stats" id="agendaCards"></div>
      </div>
      <h2 class="section-title" data-num="01">Agenda Timeline</h2>
      <div id="agendaTimeline"></div>
    </section>
    <section class="panel" id="onsite">
      <h2 class="section-title" data-num="01">Daily Agenda</h2>
      <div class="ops-grid">
        <article class="ops-card">
          <h3>Day 1 · Tuesday, September 1</h3>
          <ul class="ops-list">
            <li><div class="ops-time">9:00a-4:00p</div><div class="ops-detail"><strong>Sponsor setup / load-in</strong><span>Expo tables in Grand Ballroom 3&4.</span></div></li>
            <li><div class="ops-time">10:00a-5:00p</div><div class="ops-detail"><strong>Breakout room setup</strong><span>Classroom setup for breakout rooms.</span></div></li>
            <li><div class="ops-time">1:00-5:00p</div><div class="ops-detail"><strong>Registration open</strong><span>All-attendee registration; staff shifts are 1-3 and 3-5.</span></div></li>
            <li><div class="ops-time">5:00-7:00p</div><div class="ops-detail"><strong>Welcome Reception & Expo</strong><span>First all-attendee moment in the expo hall.</span></div></li>
            <li><div class="ops-time">7:30-9:00p</div><div class="ops-detail"><strong>VIP Dinner</strong><span>C&S Oyster House for Premium ticket holders; volunteer support starts at 7:00p.</span></div></li>
          </ul>
        </article>
        <article class="ops-card">
          <h3>Day 2 · Wednesday, September 2</h3>
          <ul class="ops-list">
            <li><div class="ops-time">8:00-9:00a</div><div class="ops-detail"><strong>Breakfast & Expo</strong><span>General Session / expo opening block.</span></div></li>
            <li><div class="ops-time">9:00a-5:00p</div><div class="ops-detail"><strong>General session + breakouts</strong><span>Keynotes, sponsor sessions, panels, Revii Showcase, and afternoon tracks.</span></div></li>
            <li><div class="ops-time">2:30-4:00p</div><div class="ops-detail"><strong>Scanner pinch point</strong><span>Three concurrent rooms need six scanners; keep Sales working expo coverage.</span></div></li>
            <li><div class="ops-time">5:00-6:30p</div><div class="ops-detail"><strong>All-attendee break</strong><span>No expo. Shuttles load from the hotel motor court starting at 6:00p.</span></div></li>
            <li><div class="ops-time">6:30-10:00p</div><div class="ops-detail"><strong>Networking Reception</strong><span>Xfinity Club at Braves Stadium, sponsored by CCH Suretax, open to all attendees.</span></div></li>
          </ul>
        </article>
        <article class="ops-card">
          <h3>Day 3 · Thursday, September 3</h3>
          <ul class="ops-list">
            <li><div class="ops-time">8:00-9:00a</div><div class="ops-detail"><strong>Breakfast keynote spotlight</strong><span>Liongard sponsor keynote spotlight.</span></div></li>
            <li><div class="ops-time">9:00a-3:15p</div><div class="ops-detail"><strong>Roadmap, workshops, breakouts</strong><span>Awards, product roadmap, AI monetization, onboarding, sales, BI, and operating model sessions.</span></div></li>
            <li><div class="ops-time">3:00-3:15p</div><div class="ops-detail"><strong>Closing remarks</strong><span>Thank-you moment and 2027 sign-up push.</span></div></li>
            <li><div class="ops-time">3:15-5:00p</div><div class="ops-detail"><strong>Teardown & pack-out</strong><span>All rooms; four-person crew noted in the handoff.</span></div></li>
          </ul>
        </article>
        <article class="ops-card">
          <h3>Volunteer / Scanner Needs</h3>
          <ul class="ops-list">
            <li><div class="ops-time">40 shifts</div><div class="ops-detail"><strong>Volunteer schedule</strong><span>76 slots across 58 staff.</span></div></li>
            <li><div class="ops-time">2 per room</div><div class="ops-detail"><strong>Breakout coverage</strong><span>Scan attendees in Vfairs, track headcount, manage speaker timecards, escalate AV issues.</span></div></li>
            <li><div class="ops-time">Tue/Wed</div><div class="ops-detail"><strong>Registration shifts</strong><span>Tue 1-3 / 3-5 and Wed 8-10 / 10-12.</span></div></li>
            <li><div class="ops-time">Receptions</div><div class="ops-detail"><strong>Floaters + F&B monitors</strong><span>Welcome Reception uses two floaters and one F&B monitor; VIP Dinner uses two volunteers.</span></div></li>
          </ul>
        </article>
      </div>
    </section>
    <section class="panel" id="links">
      <h2 class="section-title" data-num="01">Important Links</h2>
      <div class="link-grid">
        <a class="link-card" href="https://reviosummit.com/en/registration-form" target="_blank" rel="noopener">
          <h3>Registration Form</h3>
          <p>Vfairs-powered attendee registration and checkout.</p>
          <span class="link-url">reviosummit.com/en/registration-form</span>
        </a>
        <a class="link-card" href="https://book.passkey.com/e/51170051" target="_blank" rel="noopener">
          <h3>Hotel Room Block</h3>
          <p>Attendee Passkey block for Renaissance Atlanta Waverly. Cut-off: Aug 10.</p>
          <span class="link-url">book.passkey.com/e/51170051</span>
        </a>
        <a class="link-card" href="https://revio.sharepoint.com/sites/Marketing990/Shared%20Documents/Marketing%20Team/2026%20Projects/2026%20Rev.io%20Summit" target="_blank" rel="noopener">
          <h3>Volunteer Schedule</h3>
          <p>Summit2026_Volunteer_Schedule.xlsx: 40 shifts, 76 slots, registration shifts, scanner coverage, and onsite roles.</p>
          <span class="link-url">SharePoint · 2026 Rev.io Summit folder</span>
        </a>
        <a class="link-card" href="https://rev.vfairs.com" target="_blank" rel="noopener">
          <h3>Registration Exports</h3>
          <p>Pull fresh Vfairs registration reports before quoting attendee, hotel, shirt, reception, or referral counts.</p>
          <span class="link-url">Vfairs backend · RegReport export</span>
        </a>
      </div>
    </section>
    <section class="panel" id="referrals">
      <h2 class="section-title" data-num="01">Referral Contest</h2>
      <div class="metric-grid" id="referralCards"></div>
      <div class="chart-grid">
        <div class="chart-card"><h3 class="chart-title">Leaderboard</h3><div id="refLeaderboard"></div></div>
        <div class="chart-card"><h3 class="chart-title">Referral Log</h3><div class="table-wrap"><table id="referralTable"></table></div></div>
      </div>
      <h2 class="section-title" data-num="02">Comp Ticket Allocation</h2>
      <div id="discountCards" class="metric-grid"></div>
    </section>
    <section class="panel" id="sponsorships">
      <div class="subtabs">
        <button class="subtab active" data-subtab="signed">Signed Sponsors</button>
        <button class="subtab" data-subtab="outreach">Outreach Tracker</button>
        <button class="subtab" data-subtab="packages">Packages & Pricing</button>
      </div>
      <div class="subpanel active" id="signed">
        <h2 class="section-title" data-num="01">Summit Sponsorship Opportunities</h2>
        <div class="metric-grid" id="sponsorCards"></div>
        <div class="chart-card">
          <h3 class="chart-title">Salesforce Sponsor Tracker</h3>
          <div class="toolbar">
            <input id="sponsorSearch" type="search" placeholder="Search company, opportunity, package">
            <select id="sponsorStageFilter"><option value="">All stages</option></select>
            <select id="sponsorTypeFilter"><option value="">All sponsorship types</option></select>
          </div>
          <div class="table-wrap"><table id="sponsorTable"></table></div>
        </div>
      </div>
      <div class="subpanel" id="outreach">
        <h2 class="section-title" data-num="02">Outreach Tracker</h2>
        <div class="metric-grid" id="outreachCards"></div>
        <div class="table-wrap"><table id="outreachTable"></table></div>
      </div>
      <div class="subpanel" id="packages">
        <h2 class="section-title" data-num="03">Packages & Pricing</h2>
        <div class="table-wrap"><table id="packageTable"></table></div>
      </div>
    </section>
  </main>
  <footer class="footer" id="footer"></footer>
</div>
<script>window.SUMMIT_DATA = ${json};</script>
<script>
const DASHBOARD_PASSWORD = 'Revii26';
const AUTH_STORAGE_KEY = 'revioSummitDashboardAuthorized';
const authGate = document.getElementById('authGate');
const authForm = document.getElementById('authForm');
const authInput = document.getElementById('dashboardPassword');
const authError = document.getElementById('authError');

function unlockDashboard() {
  document.body.classList.remove('auth-locked');
  authGate.classList.add('hidden');
}

if (sessionStorage.getItem(AUTH_STORAGE_KEY) === 'true') {
  unlockDashboard();
} else {
  authInput.focus();
}

authForm.addEventListener('submit', (event) => {
  event.preventDefault();
  if (authInput.value === DASHBOARD_PASSWORD) {
    sessionStorage.setItem(AUTH_STORAGE_KEY, 'true');
    unlockDashboard();
    return;
  }
  authInput.value = '';
  authError.textContent = 'Incorrect password.';
  authInput.focus();
});

const data = window.SUMMIT_DATA;
const fmtMoney = (n) => (n < 0 ? '(' : '') + '$' + Math.abs(Math.round(n)).toLocaleString() + (n < 0 ? ')' : '');
const fmtPct = (n, d) => d ? Math.round((n / d) * 100) + '%' : '0%';
const takeTop = (items, limit = 8) => [...items].slice(0, limit);
const safe = (v) => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let registrantSort = { key: 'dateRegistered', direction: 'desc' };
const registrantColumnFilters = { name: '', company: '', jobTitle: '', state: '', dateRegistered: '', referral: '', packageName: '', amount: '' };

function card(label, value, note, progress) {
  const width = Math.max(0, Math.min(100, progress || 0));
  return '<article class="metric-card"><div class="metric-label">' + label + '</div><div class="metric-value">' + value + '</div><div class="metric-note">' + note + '</div><div class="bar-track"><div class="bar-fill" style="width:' + width + '%"></div></div></article>';
}

function bars(id, items, limit = 8) {
  const max = Math.max(...items.map(i => i.count), 1);
  document.getElementById(id).innerHTML = takeTop(items, limit).map(i =>
    '<div class="row-bar"><div class="row-name">' + safe(i.name) + '</div><div class="mini-track"><div class="mini-fill" style="width:' + ((i.count / max) * 100) + '%"></div></div><div class="row-count">' + i.count + '</div></div>'
  ).join('');
}

function renderHome() {
  const s = data.summary, g = data.goals;
  document.getElementById('revenueCards').innerHTML = [
    card('Ticket Revenue', fmtMoney(s.ticketRevenue), fmtPct(s.ticketRevenue, g.ticketRevenue) + ' of ' + fmtMoney(g.ticketRevenue), (s.ticketRevenue / g.ticketRevenue) * 100),
    card('Sponsorship Revenue', fmtMoney(s.sponsorRevenue), fmtPct(s.sponsorRevenue, g.sponsorRevenue) + ' of ' + fmtMoney(g.sponsorRevenue), (s.sponsorRevenue / g.sponsorRevenue) * 100)
  ].join('');
  document.getElementById('registrationCards').innerHTML = [
    card('Total Registrants', s.totalRegistrants, fmtPct(s.totalRegistrants, g.totalRegistrants) + ' of ' + g.totalRegistrants, (s.totalRegistrants / g.totalRegistrants) * 100),
    card('Paid Tickets', s.paidTickets, fmtPct(s.paidTickets, g.paidTickets) + ' of ' + g.paidTickets, (s.paidTickets / g.paidTickets) * 100),
    card('Comp Tickets', s.compTickets, s.fullCompTickets + ' full-discount tickets', Math.min(100, (s.compTickets / g.customerProspectComped) * 100)),
    card('Unique Companies', s.uniqueCompanies, 'Current registered companies', Math.min(100, (s.uniqueCompanies / 100) * 100))
  ].join('');
  renderModelTable();
  bars('businessTypes', data.breakdowns.businessTypes);
  bars('departments', data.breakdowns.departments);
  bars('jobRanks', data.breakdowns.jobRanks);
  bars('registrationsByMonth', data.breakdowns.registrationsByMonth);
}

function renderModelTable() {
  const s = data.summary, g = data.goals, a = data.actual2025;
  const rows = [
    ['Total Registrations', a.totalRegistrations, g.totalRegistrants, s.totalRegistrants, fmtPct(s.totalRegistrants, g.totalRegistrants)],
    ['Staff Tickets', a.staffTickets, g.staffTickets, s.staffTickets, fmtPct(s.staffTickets, g.staffTickets)],
    ['Customer/Prospect Registrations', a.customerProspectRegistrations, g.customerProspectRegistrations, s.customerProspectRegistrations, fmtPct(s.customerProspectRegistrations, g.customerProspectRegistrations)],
    ['Customer/Prospect Paid', a.customerProspectPaid, g.customerProspectPaid, s.paidTickets, fmtPct(s.paidTickets, g.customerProspectPaid)],
    ['Customer/Prospect Comped', a.customerProspectComped, g.customerProspectComped, s.compTickets, fmtPct(s.compTickets, g.customerProspectComped)],
    ['Avg Paid Ticket Price', fmtMoney(a.avgTicketPrice), '$448 goal', fmtMoney(s.avgPaidTicket), ''],
    ['Ticket Revenue', fmtMoney(a.ticketRevenue), fmtMoney(g.ticketRevenue), fmtMoney(s.ticketRevenue), fmtPct(s.ticketRevenue, g.ticketRevenue)],
    ['Sponsor Tickets', a.sponsorTickets, g.sponsorPasses, s.sponsorPasses, fmtPct(s.sponsorPasses, g.sponsorPasses)],
    ['# of Sponsors', a.sponsors, g.uniqueSponsors, data.sponsorships.uniqueCompanies, fmtPct(data.sponsorships.uniqueCompanies, g.uniqueSponsors)],
    ['Sponsorship Revenue', fmtMoney(a.sponsorshipRevenue), fmtMoney(g.sponsorRevenue), fmtMoney(s.sponsorRevenue), fmtPct(s.sponsorRevenue, g.sponsorRevenue)]
  ];
  document.getElementById('modelTable').innerHTML = '<thead><tr><th>Metric</th><th class="num">2025 Actuals</th><th class="num">2026 Goal</th><th class="num">2026 Current</th><th class="num">% to Goal</th></tr></thead><tbody>' +
    rows.map(r => '<tr><td><strong>' + safe(r[0]) + '</strong></td><td class="num">' + r[1] + '</td><td class="num">' + r[2] + '</td><td class="num">' + r[3] + '</td><td class="num">' + r[4] + '</td></tr>').join('') +
    '</tbody>';
}

function renderRegistrants() {
  const s = data.summary, g = data.goals;
  document.getElementById('registrantCards').innerHTML = [
    card('Total Registrants', s.totalRegistrants, fmtPct(s.totalRegistrants, g.totalRegistrants) + ' to goal', (s.totalRegistrants / g.totalRegistrants) * 100),
    card('Unique Companies', s.uniqueCompanies, 'Registered companies', Math.min(100, s.uniqueCompanies)),
    card('Ticket Revenue', fmtMoney(s.ticketRevenue), fmtPct(s.ticketRevenue, g.ticketRevenue) + ' to goal', (s.ticketRevenue / g.ticketRevenue) * 100),
    card('Hotel Rooms Needed', s.hotelRooms, 'Yes responses from registration', Math.min(100, (s.hotelRooms / s.totalRegistrants) * 100))
  ].join('');
  const ticketSelect = document.getElementById('ticketFilter');
  ticketSelect.innerHTML = '<option value="">All ticket types</option>' + data.breakdowns.ticketTypes.map(i => '<option>' + safe(i.name) + '</option>').join('');
  const companies = [...new Set(data.registrants.map(r => r.company))].sort((a,b) => a.localeCompare(b));
  document.getElementById('companyFilter').innerHTML = '<option value="">All companies</option>' + companies.map(c => '<option>' + safe(c) + '</option>').join('');
  const referrers = [...new Set(data.registrants.map(r => r.referral).filter(Boolean))].sort((a,b) => a.localeCompare(b));
  document.getElementById('referrerFilter').innerHTML = '<option value="">All referrers</option>' + referrers.map(r => '<option>' + safe(r) + '</option>').join('');
  document.getElementById('search').addEventListener('input', renderRegistrantTable);
  ticketSelect.addEventListener('change', renderRegistrantTable);
  document.getElementById('companyFilter').addEventListener('change', renderRegistrantTable);
  document.getElementById('referrerFilter').addEventListener('change', renderRegistrantTable);
  renderRegistrantTable();
  renderTicketTable();
}

function uniqueRegistrantValues(key) {
  return [...new Set(data.registrants.map(r => r[key]).filter(Boolean))].sort((a,b) => String(a).localeCompare(String(b)));
}

function registrantMonthOptions() {
  const byValue = new Map();
  data.registrants.forEach(r => {
    if (!r.dateRegistered) return;
    const value = r.dateRegistered.slice(0, 7);
    const label = new Date(value + '-01T00:00:00Z').toLocaleDateString('en-US', { month: 'short', year: 'numeric', timeZone: 'UTC' });
    byValue.set(value, label);
  });
  return [...byValue.entries()].sort((a,b) => b[0].localeCompare(a[0]));
}

function optionHtml(items, selectedValue = '') {
  return items.map(item => {
    const value = Array.isArray(item) ? item[0] : item;
    const label = Array.isArray(item) ? item[1] : item;
    return '<option value="' + safe(value) + '"' + (String(value) === String(selectedValue) ? ' selected' : '') + '>' + safe(label) + '</option>';
  }).join('');
}

function columnInput(key, placeholder) {
  return '<input class="column-filter" data-col-filter="' + key + '" value="' + safe(registrantColumnFilters[key]) + '" placeholder="' + safe(placeholder) + '">';
}

function columnSelect(key, placeholder, options) {
  return '<select class="column-filter" data-col-filter="' + key + '"><option value="">' + safe(placeholder) + '</option>' + optionHtml(options, registrantColumnFilters[key]) + '</select>';
}

function amountBucket(row) {
  return row.amount > 0 ? 'Paid' : 'Comp';
}

function sortValue(row, key) {
  if (key === 'amount') return Number(row.amount || 0);
  if (key === 'dateRegistered') return row.dateRegistered || '';
  return String(row[key] || '').toLowerCase();
}

function sortRows(rows) {
  const { key, direction } = registrantSort;
  const dir = direction === 'asc' ? 1 : -1;
  return [...rows].sort((a, b) => {
    const av = sortValue(a, key), bv = sortValue(b, key);
    if (av < bv) return -1 * dir;
    if (av > bv) return 1 * dir;
    return String(a.name || '').localeCompare(String(b.name || ''));
  });
}

function sortHeader(label, key, cls = '') {
  const active = registrantSort.key === key;
  const direction = active ? registrantSort.direction : '';
  return '<th' + (cls ? ' class="' + cls + '"' : '') + '><button class="sort-button ' + (active ? 'active ' + direction : '') + '" type="button" data-sort-key="' + key + '">' + safe(label) + '</button></th>';
}

function renderRegistrantTable(focusColumn = '') {
  const q = document.getElementById('search').value.toLowerCase();
  const ticket = document.getElementById('ticketFilter').value;
  const company = document.getElementById('companyFilter').value;
  const referrer = document.getElementById('referrerFilter').value;
  const rows = sortRows(data.registrants.filter(r => {
    const hay = [r.name, r.company, r.jobTitle, r.discountCode, r.referral, r.state, r.dateLabel, r.packageName, amountBucket(r)].join(' ').toLowerCase();
    const f = registrantColumnFilters;
    return (!q || hay.includes(q))
      && (!ticket || r.packageName === ticket)
      && (!company || r.company === company)
      && (!referrer || r.referral === referrer)
      && (!f.name || String(r.name || '').toLowerCase().includes(f.name.toLowerCase()))
      && (!f.company || r.company === f.company)
      && (!f.jobTitle || String(r.jobTitle || '').toLowerCase().includes(f.jobTitle.toLowerCase()))
      && (!f.state || r.state === f.state)
      && (!f.dateRegistered || String(r.dateRegistered || '').startsWith(f.dateRegistered))
      && (!f.referral || r.referral === f.referral)
      && (!f.packageName || r.packageName === f.packageName)
      && (!f.amount || amountBucket(r) === f.amount);
  }));
  renderActiveRegistrantFilter(rows.length);
  document.getElementById('registrantsTable').innerHTML =
    '<thead><tr>' +
    sortHeader('Name', 'name') + sortHeader('Company', 'company') + sortHeader('Title', 'jobTitle') + sortHeader('State', 'state') + sortHeader('Registered', 'dateRegistered') + sortHeader('Referred By', 'referral') + sortHeader('Ticket', 'packageName') + sortHeader('Amount', 'amount', 'num') +
    '</tr><tr class="filter-row"><th>' + columnInput('name', 'Filter name') + '</th><th>' + columnSelect('company', 'All', uniqueRegistrantValues('company')) + '</th><th>' + columnInput('jobTitle', 'Filter title') + '</th><th>' + columnSelect('state', 'All', uniqueRegistrantValues('state')) + '</th><th>' + columnSelect('dateRegistered', 'All months', registrantMonthOptions()) + '</th><th>' + columnSelect('referral', 'All', uniqueRegistrantValues('referral')) + '</th><th>' + columnSelect('packageName', 'All', uniqueRegistrantValues('packageName')) + '</th><th>' + columnSelect('amount', 'All', ['Paid', 'Comp']) + '</th></tr></thead><tbody>' +
    (rows.length ? rows.map(r => '<tr><td><strong>' + safe(r.name) + '</strong></td><td>' + safe(r.company) + '</td><td>' + safe(r.jobTitle || '--') + '</td><td>' + safe(r.state || '--') + '</td><td>' + safe(r.dateLabel || '--') + '</td><td>' + (r.referral ? '<button class="pill click-row" type="button" data-referrer="' + safe(r.referral) + '">' + safe(r.referral) + '</button>' : '<span class="muted">--</span>') + '</td><td>' + safe(r.packageName) + (r.discountCode ? '<div class="muted">' + safe(r.discountCode) + '</div>' : '') + '</td><td class="num">' + (r.amount ? fmtMoney(r.amount) : '<span class="pill">Comp</span>') + '</td></tr>').join('') : '<tr><td colspan="8" class="muted">No registrants match the current filters.</td></tr>') +
    '</tbody>';
  document.querySelectorAll('#registrantsTable [data-referrer]').forEach(btn => btn.addEventListener('click', () => applyReferralFilter(btn.dataset.referrer)));
  document.querySelectorAll('#registrantsTable [data-sort-key]').forEach(btn => btn.addEventListener('click', () => {
    const key = btn.dataset.sortKey;
    registrantSort = registrantSort.key === key ? { key, direction: registrantSort.direction === 'asc' ? 'desc' : 'asc' } : { key, direction: key === 'dateRegistered' || key === 'amount' ? 'desc' : 'asc' };
    renderRegistrantTable();
  }));
  document.querySelectorAll('#registrantsTable [data-col-filter]').forEach(control => {
    control.addEventListener('input', () => { registrantColumnFilters[control.dataset.colFilter] = control.value; renderRegistrantTable(control.dataset.colFilter); });
    control.addEventListener('change', () => { registrantColumnFilters[control.dataset.colFilter] = control.value; renderRegistrantTable(control.dataset.colFilter); });
  });
  if (focusColumn) {
    const activeFilter = document.querySelector('#registrantsTable [data-col-filter="' + focusColumn + '"]');
    if (activeFilter) {
      activeFilter.focus();
      if (activeFilter.setSelectionRange && activeFilter.tagName === 'INPUT') {
        const pos = activeFilter.value.length;
        activeFilter.setSelectionRange(pos, pos);
      }
    }
  }
}

function renderActiveRegistrantFilter(count) {
  const referrer = document.getElementById('referrerFilter').value;
  const target = document.getElementById('activeRegistrantFilter');
  if (!referrer) {
    target.classList.remove('visible');
    target.innerHTML = '';
    return;
  }
  target.classList.add('visible');
  target.innerHTML = '<span class="pill">' + count + ' registrants referred by ' + safe(referrer) + '</span><button class="action-button" type="button" id="clearRegistrantFilters">Clear referral filter</button>';
  document.getElementById('clearRegistrantFilters').addEventListener('click', () => applyReferralFilter(''));
}

function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tabName));
  document.querySelectorAll('.panel').forEach(p => p.classList.toggle('active', p.id === tabName));
}

function applyReferralFilter(referrer) {
  switchTab('registrants');
  document.getElementById('referrerFilter').value = referrer;
  document.getElementById('search').value = '';
  document.getElementById('ticketFilter').value = '';
  document.getElementById('companyFilter').value = '';
  Object.keys(registrantColumnFilters).forEach(key => { registrantColumnFilters[key] = ''; });
  renderRegistrantTable();
  document.getElementById('registrants').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderTicketTable() {
  document.getElementById('ticketTable').innerHTML =
    '<thead><tr><th>Ticket Type</th><th class="num">Sold</th><th class="num">Paid</th><th class="num">Comp</th><th class="num">Revenue</th></tr></thead><tbody>' +
    data.breakdowns.ticketTypes.map(r => '<tr><td><strong>' + safe(r.name) + '</strong></td><td class="num">' + r.count + '</td><td class="num">' + r.paid + '</td><td class="num">' + r.comp + '</td><td class="num">' + fmtMoney(r.revenue) + '</td></tr>').join('') +
    '</tbody>';
}

function renderAgenda() {
  const days = data.agenda || [];
  const sessions = days.flatMap(day => day.sessions);
  const breakoutCount = sessions.filter(s => s.type === 'Breakout' || s.type === 'Workshop').length;
  const networkingCount = sessions.filter(s => /Networking|Meal|Expo|Break/i.test(s.type + ' ' + s.title)).length;
  document.getElementById('agendaCards').innerHTML = [
    card('Program Days', days.length, 'September 1-3, 2026', 100),
    card('Sessions + Events', sessions.length, 'Agenda blocks currently listed', 100),
    card('Breakouts + Workshops', breakoutCount, 'Track-based learning sessions', Math.min(100, breakoutCount * 8)),
    card('Networking Moments', networkingCount, 'Meals, expo, breaks, receptions', Math.min(100, networkingCount * 6))
  ].join('');
  document.getElementById('agendaTimeline').innerHTML = days.map(day =>
    '<article class="agenda-day"><div class="agenda-day-header"><div><div class="agenda-day-label">' + safe(day.label) + '</div><div class="agenda-day-title">' + safe(day.date) + '</div></div><div class="agenda-day-theme">' + safe(day.theme) + '</div></div><div class="agenda-list">' +
    day.sessions.map(session => {
      const meta = [session.type, session.track].filter(Boolean).join(' / ');
      return '<div class="agenda-item"><div class="agenda-time">' + safe(session.start) + ' - ' + safe(session.end) + '</div><div><div class="agenda-title">' + safe(session.title) + '</div>' + (meta ? '<div class="agenda-meta">' + safe(meta) + '</div>' : '') + '</div><div class="agenda-meta">' + safe(session.track || '') + '</div><div class="agenda-room">' + safe(session.room || '') + '</div></div>';
    }).join('') +
    '</div></article>'
  ).join('');
}

function renderReferrals() {
  const refs = data.referrals;
  const leaderboard = {};
  refs.forEach(r => leaderboard[r.referrer] = (leaderboard[r.referrer] || 0) + 1);
  const ranked = Object.entries(leaderboard).map(([name, count]) => ({ name, count })).sort((a,b) => b.count - a.count || a.name.localeCompare(b.name));
  document.getElementById('referralCards').innerHTML = [
    card('Total Referrals', refs.length, 'Registrant referral field', Math.min(100, refs.length * 2)),
    card('Unique Referrers', ranked.length, 'People credited', Math.min(100, ranked.length * 4)),
    card('Current Leader', ranked[0] ? safe(ranked[0].name) : '--', ranked[0] ? ranked[0].count + ' referrals' : 'No referrals yet', ranked[0] ? 100 : 0),
    card('Contest Deadline', 'Aug 21', 'Grand prize cutoff', 100)
  ].join('');
  document.getElementById('refLeaderboard').innerHTML = ranked.map((r, i) => '<button class="row-bar click-row" type="button" data-referrer="' + safe(r.name) + '" title="Filter registrants referred by ' + safe(r.name) + '"><div class="row-name">' + (i + 1) + '. ' + safe(r.name) + '</div><div class="mini-track"><div class="mini-fill" style="width:' + ((r.count / Math.max(1, ranked[0]?.count || 1)) * 100) + '%"></div></div><div class="row-count">' + r.count + '</div></button>').join('');
  document.getElementById('referralTable').innerHTML = '<thead><tr><th>Registrant</th><th>Referred By</th><th class="num">Date</th></tr></thead><tbody>' + refs.map(r => '<tr><td><strong>' + safe(r.registrant) + '</strong></td><td><button class="pill click-row" type="button" data-referrer="' + safe(r.referrer) + '">' + safe(r.referrer) + '</button></td><td class="num">' + safe(r.date) + '</td></tr>').join('') + '</tbody>';
  document.querySelectorAll('#referrals [data-referrer]').forEach(btn => btn.addEventListener('click', () => applyReferralFilter(btn.dataset.referrer)));
  document.getElementById('discountCards').innerHTML = data.breakdowns.discounts.map(d => card(d.name, d.count, d.cap == null ? 'No cap set' : d.remaining + ' remaining of ' + d.cap, d.cap == null ? 100 : (d.count / d.cap) * 100)).join('');
}

const PACKAGES = [
  ['Title Sponsor', 45000, 1, 10], ['Diamond Sponsor', 26000, 3, 6], ['Platinum Sponsor', 18000, 3, 4],
  ['Gold Sponsor', 12000, 5, 3], ['Silver Sponsor', 9000, 4, 2], ['Bronze Sponsor', 6500, 4, 1],
  ['Rookie Sponsor', 4500, 6, 1], ['Registration Sponsor', 7500, 2, 1], ['Mobile App Sponsor', 6500, 1, 1],
  ['Custom / Barter', 0, 99, 0]
];
const DEFAULT_OUTREACH = [
  ['Acronis','Title','Usman','Contacted','Platinum 2025'], ['Alianza','Rookie+','Usman','Contacted','Declined 2025'],
  ['Altaworx','Emerald','Usman','Not Contacted','Gold 2025'], ['Crexendo','Learning Track','Christina','Meeting Scheduled','Platinum 2025'],
  ['Pax8','Diamond Swap','Brook/Megan','Contacted',''], ['TaxConnex','Tier 2','Christina','Meeting Scheduled','Gold 2025'],
  ['TRX Services','Platinum','Christina','Meeting Scheduled','Platinum 2025'], ['Wolters Kluwer','Diamond','Usman','Contacted','Diamond 2025'],
  ['Cynomi','','Jake M.','Contacted',''], ['Pia','','Jake M.','Contacted','Sent prospectus']
];

function renderSponsorships() {
  const sponsorships = data.sponsorships;
  const rows = filteredSponsorRows();
  document.getElementById('sponsorCards').innerHTML = [
    card('Sponsor Opp Amount', fmtMoney(sponsorships.totalAmount), fmtPct(sponsorships.totalAmount, data.goals.sponsorRevenue) + ' to sponsor goal', (sponsorships.totalAmount / data.goals.sponsorRevenue) * 100),
    card('Closed Won Amount', fmtMoney(sponsorships.wonAmount), sponsorships.closedWon + ' closed won opps', (sponsorships.wonAmount / data.goals.sponsorRevenue) * 100),
    card('Open Pipeline Amount', fmtMoney(sponsorships.openAmount), 'Open Summit Sponsorship opps', (sponsorships.openAmount / data.goals.sponsorRevenue) * 100),
    card('Sponsor Opps', sponsorships.totalOpportunities, sponsorships.uniqueCompanies + ' unique companies', Math.min(100, (sponsorships.uniqueCompanies / data.goals.uniqueSponsors) * 100))
  ].join('');
  document.getElementById('sponsorTable').innerHTML = '<thead><tr><th>Company</th><th>Sponsorship Type</th><th>Package</th><th>Stage</th><th class="num">Close Date</th><th class="num">Amount</th></tr></thead><tbody>' +
    (rows.length ? rows.map(r => '<tr><td><strong>' + safe(r.company) + '</strong><div class="muted">' + safe(r.opportunityName) + '</div></td><td>' + safe(r.sponsorshipType) + '</td><td>' + safe(r.sponsorshipPackage || '--') + '</td><td><span class="pill">' + safe(r.stage) + '</span></td><td class="num">' + safe(r.closeDate || '--') + '</td><td class="num">' + fmtMoney(r.amount) + '</td></tr>').join('') : '<tr><td colspan="6" class="muted">No Salesforce Summit Sponsorship opportunities found.</td></tr>') +
    '</tbody>';
}
function filteredSponsorRows() {
  const search = document.getElementById('sponsorSearch')?.value.toLowerCase() || '';
  const stage = document.getElementById('sponsorStageFilter')?.value || '';
  const type = document.getElementById('sponsorTypeFilter')?.value || '';
  return data.sponsorships.opportunities.filter(r => {
    const hay = [r.company, r.opportunityName, r.sponsorshipType, r.sponsorshipPackage, r.stage].join(' ').toLowerCase();
    return (!search || hay.includes(search)) && (!stage || r.stage === stage) && (!type || r.sponsorshipType === type);
  });
}
function setupSponsorFilters() {
  const stages = [...new Set(data.sponsorships.opportunities.map(r => r.stage))].sort((a,b) => a.localeCompare(b));
  const types = [...new Set(data.sponsorships.opportunities.map(r => r.sponsorshipType).filter(t => t && t !== 'Unspecified'))].sort((a,b) => a.localeCompare(b));
  document.getElementById('sponsorStageFilter').innerHTML = '<option value="">All stages</option>' + stages.map(v => '<option>' + safe(v) + '</option>').join('');
  document.getElementById('sponsorTypeFilter').innerHTML = '<option value="">All sponsorship types</option>' + types.map(v => '<option>' + safe(v) + '</option>').join('');
  document.getElementById('sponsorSearch').addEventListener('input', renderSponsorships);
  document.getElementById('sponsorStageFilter').addEventListener('change', renderSponsorships);
  document.getElementById('sponsorTypeFilter').addEventListener('change', renderSponsorships);
}
function renderOutreachAndPackages() {
  const counts = DEFAULT_OUTREACH.reduce((acc, r) => (acc[r[3]] = (acc[r[3]] || 0) + 1, acc), {});
  document.getElementById('outreachCards').innerHTML = ['Not Contacted','Contacted','Meeting Scheduled','Pending Contract'].map(k => card(k, counts[k] || 0, 'Prospects in tracker', Math.min(100, (counts[k] || 0) * 10))).join('');
  document.getElementById('outreachTable').innerHTML = '<thead><tr><th>Company</th><th>Offer</th><th>Owner</th><th>Status</th><th>2025 Context</th></tr></thead><tbody>' + DEFAULT_OUTREACH.map(r => '<tr><td><strong>' + safe(r[0]) + '</strong></td><td>' + safe(r[1] || '--') + '</td><td>' + safe(r[2]) + '</td><td><span class="pill">' + safe(r[3]) + '</span></td><td>' + safe(r[4] || '--') + '</td></tr>').join('') + '</tbody>';
  document.getElementById('packageTable').innerHTML = '<thead><tr><th>Package</th><th class="num">Price</th><th class="num">Qty Available</th><th class="num">Passes</th><th class="num">Potential</th></tr></thead><tbody>' + PACKAGES.filter(p => p[0] !== 'Custom / Barter').map(p => '<tr><td><strong>' + safe(p[0]) + '</strong></td><td class="num">' + fmtMoney(p[1]) + '</td><td class="num">' + p[2] + '</td><td class="num">' + p[3] + '</td><td class="num">' + fmtMoney(p[1] * p[2]) + '</td></tr>').join('') + '</tbody>';
}

document.querySelectorAll('.tab').forEach(btn => btn.addEventListener('click', () => {
  switchTab(btn.dataset.tab);
}));
document.querySelectorAll('.subtab').forEach(btn => btn.addEventListener('click', () => {
  document.querySelectorAll('.subtab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.subpanel').forEach(p => p.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById(btn.dataset.subtab).classList.add('active');
}));

renderHome();
renderRegistrants();
renderAgenda();
renderReferrals();
setupSponsorFilters();
renderSponsorships();
renderOutreachAndPackages();
document.getElementById('footer').textContent = 'Last updated: ' + new Date(data.generatedAt).toLocaleString('en-US', { dateStyle: 'medium', timeStyle: 'short' }) + ' · Sources: ' + data.sourceFile + ' and Salesforce Opportunity Type = ' + data.sponsorships.sourceOpportunityType + ' · Emails, phone numbers, and source user IDs excluded from dashboard build.';
</script>
</body>
</html>`;
}

async function main() {
  fs.mkdirSync(dataDir, { recursive: true });
  const rows = await loadWorkbookRows();
  const sponsorshipOpportunities = await loadSponsorshipOpportunities();
  const model = buildModel(rows, sponsorshipOpportunities);
  fs.writeFileSync(path.join(dataDir, 'summit-data.json'), JSON.stringify(model, null, 2));
  fs.writeFileSync(path.join(outDir, 'index.html'), renderDashboard(model));
  console.log(`Built dashboard with ${model.summary.totalRegistrants} registrants, ${model.summary.uniqueCompanies} companies, ${money(model.summary.ticketRevenue)} ticket revenue, and ${model.sponsorships.totalOpportunities} sponsor opps totaling ${money(model.sponsorships.totalAmount)}.`);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
