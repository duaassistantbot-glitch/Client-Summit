const fs = require('fs');
const https = require('https');
const path = require('path');
const { createRequire } = require('module');

const outDir = path.resolve(__dirname, '..');
const workspaceDir = path.resolve(outDir, '..');
const workspaceRequire = createRequire(path.join(workspaceDir, 'gmail-tool', 'package.json'));
const XLSX = workspaceRequire('xlsx');

const API_VERSION = 'v59.0';
const CREATED_SINCE = process.env.SUMMIT_OPP_CREATED_SINCE || '2026-09-01T00:00:00Z';
const workbookPath = process.env.SUMMIT_XLSX || path.join(outDir, 'assets', 'source', 'registration-report-deduped-2026-08-12.xlsx');
const dataDir = path.join(outDir, 'assets', 'data');
const jsonOut = path.join(dataDir, 'summit-opportunity-data.json');
const csvOut = path.join(outDir, 'summit-opportunities.csv');
const htmlOut = path.join(outDir, 'opportunity-dashboard.html');

function loadEnv() {
  const envPath = path.join(workspaceDir, '.env');
  if (!fs.existsSync(envPath)) return;
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const i = line.indexOf('=');
    if (i > 0) {
      const key = line.slice(0, i).trim();
      const value = line.slice(i + 1).trim();
      if (key && value && !process.env[key]) process.env[key] = value;
    }
  }
}

function sfRequest({ method = 'GET', requestPath, token, body, headers = {} }) {
  const instanceUrl = process.env.SF_INSTANCE_URL;
  if (!instanceUrl) throw new Error('SF_INSTANCE_URL is not configured.');
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: new URL(instanceUrl).hostname,
      path: requestPath,
      method,
      headers: { ...(token ? { Authorization: `Bearer ${token}` } : {}), ...headers }
    }, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        let parsed;
        try { parsed = data ? JSON.parse(data) : null; } catch { return reject(new Error(data.slice(0, 1000))); }
        if (res.statusCode >= 400) {
          const msg = Array.isArray(parsed) ? parsed.map(e => e.message || e.errorCode).join('; ') : (parsed?.message || parsed?.error_description || JSON.stringify(parsed));
          return reject(new Error(msg));
        }
        resolve(parsed);
      });
    });
    req.on('error', reject);
    if (body) req.write(body);
    req.end();
  });
}

async function getSalesforceToken() {
  loadEnv();
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: process.env.SF_CLIENT_ID,
    client_secret: process.env.SF_CLIENT_SECRET
  }).toString();
  const parsed = await sfRequest({
    method: 'POST',
    requestPath: '/services/oauth2/token',
    body,
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
  });
  return parsed.access_token;
}

async function sfQueryAll(token, soql) {
  let result = await sfRequest({ token, requestPath: `/services/data/${API_VERSION}/query?q=${encodeURIComponent(soql)}` });
  let records = result.records || [];
  while (result.nextRecordsUrl) {
    result = await sfRequest({ token, requestPath: result.nextRecordsUrl });
    records = records.concat(result.records || []);
  }
  return records;
}

const q = (s) => `'${String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'")}'`;
const chunks = (arr, n = 80) => Array.from({ length: Math.ceil(arr.length / n) }, (_, i) => arr.slice(i * n, i * n + n));
const normEmail = (s) => String(s || '').trim().toLowerCase();
const normName = (s) => String(s || '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, ' ').trim().replace(/\s+/g, ' ');
const clean = (s) => String(s || '').trim();
const money = (n) => Number(n || 0).toLocaleString('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 });
const csvEscape = (v) => {
  const s = v == null ? '' : String(v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};
const escapeHtml = (value) => String(value ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));

function loadRegistrants() {
  const wb = XLSX.readFile(workbookPath, { raw: false });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
  const seen = new Set();
  return rows
    .filter(r => !/cancel/i.test(String(r.Status || '')))
    .map((r, index) => {
      const email = normEmail(r['Business Email'] || r.Email);
      const firstName = clean(r['First Name (1557966)']);
      const lastName = clean(r['Last Name (1557967)']);
      const referralFirst = clean(r['Referral First Name']);
      const referralLast = clean(r['Referral Last Name']);
      return {
        sourceRow: index + 2,
        email,
        name: [firstName, lastName].filter(Boolean).join(' ') || clean(r.Email),
        company: clean(r.Company),
        referral: [referralFirst, referralLast].filter(Boolean).join(' '),
        discountCode: clean(r['Discount Code'])
      };
    })
    .filter(r => {
      const key = r.email || `${normName(r.name)}|${normName(r.company)}`;
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

async function fetchContactsAndAccounts(token, registrants) {
  const emails = [...new Set(registrants.map(r => r.email).filter(Boolean))];
  const contactByEmail = new Map();
  const contactFields = 'Id, Name, FirstName, LastName, Email, AccountId, Account.Id, Account.Name, Account.Type, Owner.Name';
  for (const part of chunks(emails, 90)) {
    const recs = await sfQueryAll(token, `SELECT ${contactFields} FROM Contact WHERE Email IN (${part.map(q).join(',')})`);
    for (const c of recs) if (c.Email && !contactByEmail.has(normEmail(c.Email))) contactByEmail.set(normEmail(c.Email), c);
  }

  const unmatchedCompanies = [...new Set(registrants.filter(r => !contactByEmail.has(r.email)).map(r => r.company).filter(Boolean))];
  const accountByNormName = new Map();
  const accountFields = 'Id, Name, Type, Owner.Name';
  for (const part of chunks(unmatchedCompanies, 90)) {
    const recs = await sfQueryAll(token, `SELECT ${accountFields} FROM Account WHERE Name IN (${part.map(q).join(',')})`);
    for (const a of recs) accountByNormName.set(normName(a.Name), a);
  }

  return registrants.map((r) => {
    const contact = contactByEmail.get(r.email) || null;
    const account = contact?.Account || accountByNormName.get(normName(r.company)) || null;
    return {
      ...r,
      sfContactId: contact?.Id || '',
      sfContactName: contact?.Name || '',
      sfAccountId: account?.Id || '',
      sfAccountName: account?.Name || '',
      sfAccountType: account?.Type || '',
      sfContactOwner: contact?.Owner?.Name || '',
      matchMethod: contact?.Account ? 'Contact email → Account' : account ? 'Registrant company → Account exact name' : contact ? 'Contact email only' : 'Unmatched'
    };
  });
}

async function fetchOpportunities(token, enrichedRegistrants) {
  const contactIds = [...new Set(enrichedRegistrants.map(r => r.sfContactId).filter(Boolean))];
  const accountIds = [...new Set(enrichedRegistrants.map(r => r.sfAccountId).filter(Boolean))];
  const oppFields = 'Id, Name, AccountId, Account.Name, Type, Product_Type__c, Total_Monthly_Fees__c, Renewal_Amount__c, Amount, StageName, IsWon, IsClosed, CreatedDate, CloseDate, Owner.Name';

  const byContact = [];
  for (const part of chunks(contactIds, 90)) {
    byContact.push(...await sfQueryAll(token, `
      SELECT ContactId, OpportunityId, ${oppFields.split(', ').map(f => `Opportunity.${f}`).join(', ')}
      FROM OpportunityContactRole
      WHERE ContactId IN (${part.map(q).join(',')})
        AND Opportunity.CreatedDate >= ${CREATED_SINCE}
    `));
  }

  const byAccount = [];
  for (const part of chunks(accountIds, 90)) {
    byAccount.push(...await sfQueryAll(token, `SELECT ${oppFields} FROM Opportunity WHERE AccountId IN (${part.map(q).join(',')}) AND CreatedDate >= ${CREATED_SINCE} ORDER BY CreatedDate DESC`));
  }

  return { byContact, byAccount };
}

function buildRows(enrichedRegistrants, opportunities) {
  const registrantsByContact = new Map();
  const registrantsByAccount = new Map();
  for (const r of enrichedRegistrants) {
    if (r.sfContactId) {
      const arr = registrantsByContact.get(r.sfContactId) || [];
      arr.push(r);
      registrantsByContact.set(r.sfContactId, arr);
    }
    if (r.sfAccountId) {
      const arr = registrantsByAccount.get(r.sfAccountId) || [];
      arr.push(r);
      registrantsByAccount.set(r.sfAccountId, arr);
    }
  }

  const rowsByKey = new Map();
  const add = (opp, registrant, connectionType) => {
    if (!opp || !registrant) return;
    const key = opp.Id;
    if (rowsByKey.has(key)) {
      const existing = rowsByKey.get(key);
      if (connectionType === 'Opportunity Contact Role' && existing.connectionType !== 'Opportunity Contact Role') {
        existing.attendee = registrant.sfContactName || registrant.name || '';
        existing.referredBy = registrant.referral || 'Not provided';
        existing.connectionType = connectionType;
        existing.sfContactId = registrant.sfContactId || '';
        existing.primaryContactOwner = registrant.sfContactOwner || '';
        existing.sourceRegistrantRow = registrant.sourceRow || '';
      }
      return;
    }
    const mrr = Number(opp.Total_Monthly_Fees__c || opp.Renewal_Amount__c || opp.Amount || 0) || 0;
    rowsByKey.set(key, {
      opportunityId: opp.Id,
      opportunityName: opp.Name || '',
      account: opp.Account?.Name || registrant.sfAccountName || '',
      accountId: opp.AccountId || registrant.sfAccountId || '',
      productType: opp.Product_Type__c || opp.Type || 'Unspecified',
      mrr,
      totalMonthlyFees: Number(opp.Total_Monthly_Fees__c || 0) || 0,
      renewalAmount: Number(opp.Renewal_Amount__c || 0) || 0,
      amount: Number(opp.Amount || 0) || 0,
      stage: opp.StageName || '',
      createdDate: opp.CreatedDate || '',
      closeDate: opp.CloseDate || '',
      opportunityOwner: opp.Owner?.Name || '',
      attendee: registrant.sfContactName || registrant.name || '',
      primaryContactOwner: registrant.sfContactOwner || '',
      attendeeCompanyFromRegistration: registrant.company || '',
      referredBy: registrant.referral || 'Not provided',
      connectionType,
      registrantMatchMethod: registrant.matchMethod || '',
      sfContactId: registrant.sfContactId || '',
      sourceRegistrantRow: registrant.sourceRow || ''
    });
  };

  for (const role of opportunities.byContact) {
    const regs = registrantsByContact.get(role.ContactId) || [];
    for (const r of regs) add({ ...role.Opportunity, Id: role.OpportunityId }, r, 'Opportunity Contact Role');
  }

  for (const opp of opportunities.byAccount) {
    const regs = registrantsByAccount.get(opp.AccountId) || [];
    for (const r of regs) add(opp, r, opp.Owner?.Name && r.sfContactId ? 'Attendee Account' : 'Registrant Account');
  }

  return [...rowsByKey.values()].sort((a, b) => {
    const dateCompare = String(b.createdDate).localeCompare(String(a.createdDate));
    if (dateCompare) return dateCompare;
    return a.account.localeCompare(b.account) || a.attendee.localeCompare(b.attendee);
  });
}

function summarize(rows, enrichedRegistrants) {
  const uniqueOpps = new Map();
  for (const row of rows) {
    if (!uniqueOpps.has(row.opportunityId)) uniqueOpps.set(row.opportunityId, row);
  }
  const opps = [...uniqueOpps.values()];
  const productMap = new Map();
  for (const opp of opps) {
    const key = opp.productType || 'Unspecified';
    const current = productMap.get(key) || { productType: key, opportunities: 0, mrr: 0, amount: 0 };
    current.opportunities += 1;
    current.mrr += opp.mrr;
    current.amount += opp.amount;
    productMap.set(key, current);
  }
  return {
    generatedAt: new Date().toISOString(),
    sourceFile: path.relative(outDir, workbookPath),
    createdSince: CREATED_SINCE,
    registrants: enrichedRegistrants.length,
    matchedRegistrants: enrichedRegistrants.filter(r => r.sfAccountId || r.sfContactId).length,
    uniqueMatchedAccounts: new Set(enrichedRegistrants.map(r => r.sfAccountId).filter(Boolean)).size,
    opportunityRows: rows.length,
    uniqueOpportunities: opps.length,
    uniqueAccountsWithOpportunities: new Set(opps.map(o => o.accountId).filter(Boolean)).size,
    totalMRR: opps.reduce((sum, row) => sum + row.mrr, 0),
    totalAmount: opps.reduce((sum, row) => sum + row.amount, 0),
    wonOpportunities: opps.filter(o => /closed won/i.test(o.stage)).length,
    openOpportunities: opps.filter(o => !/^closed/i.test(o.stage)).length,
    byProductType: [...productMap.values()].sort((a, b) => b.mrr - a.mrr || b.opportunities - a.opportunities)
  };
}

function writeCsv(rows) {
  const headers = ['Account', 'Opportunity Product Type', 'MRR', 'Contact Who Attended Summit', 'Who Referred Attendee', 'Opportunity Owner', 'Primary Contact Owner', 'Opportunity', 'Stage', 'Created Date', 'Close Date', 'Connection Type', 'Amount', 'Salesforce Account ID', 'Salesforce Contact ID', 'Salesforce Opportunity ID'];
  const data = rows.map(r => ({
    'Account': r.account,
    'Opportunity Product Type': r.productType,
    'MRR': r.mrr,
    'Contact Who Attended Summit': r.attendee,
    'Who Referred Attendee': r.referredBy,
    'Opportunity Owner': r.opportunityOwner,
    'Primary Contact Owner': r.primaryContactOwner,
    'Opportunity': r.opportunityName,
    'Stage': r.stage,
    'Created Date': r.createdDate,
    'Close Date': r.closeDate,
    'Connection Type': r.connectionType,
    'Amount': r.amount,
    'Salesforce Account ID': r.accountId,
    'Salesforce Contact ID': r.sfContactId,
    'Salesforce Opportunity ID': r.opportunityId
  }));
  fs.writeFileSync(csvOut, [headers.join(','), ...data.map(row => headers.map(h => csvEscape(row[h])).join(','))].join('\n'));
}

function renderHtml(summary, rows) {
  const productRows = summary.byProductType.map(p => `<tr><td>${escapeHtml(p.productType)}</td><td>${p.opportunities}</td><td>${money(p.mrr)}</td><td>${money(p.amount)}</td></tr>`).join('');
  const rowHtml = rows.map(r => `<tr data-product="${escapeHtml(r.productType)}" data-owner="${escapeHtml(r.opportunityOwner)}" data-stage="${escapeHtml(r.stage)}" data-search="${escapeHtml([r.account, r.productType, r.attendee, r.referredBy, r.opportunityOwner, r.opportunityName, r.stage, r.connectionType].join(' ').toLowerCase())}">
    <td><strong>${escapeHtml(r.account)}</strong><span>${escapeHtml(r.opportunityName)}</span></td>
    <td>${escapeHtml(r.productType)}</td>
    <td data-num="${r.mrr}">${money(r.mrr)}</td>
    <td>${escapeHtml(r.attendee)}<span>Contact owner: ${escapeHtml(r.primaryContactOwner || 'Unassigned')}</span></td>
    <td>${escapeHtml(r.referredBy)}</td>
    <td>${escapeHtml(r.opportunityOwner)}</td>
    <td>${escapeHtml(r.stage)}<span>${escapeHtml((r.createdDate || '').slice(0, 10))}</span></td>
  </tr>`).join('');
  const products = [...new Set(rows.map(r => r.productType || 'Unspecified'))].sort();
  const owners = [...new Set(rows.map(r => r.opportunityOwner || 'Unassigned'))].sort();
  const stages = [...new Set(rows.map(r => r.stage || 'Unspecified'))].sort();
  const options = (items) => items.map(x => `<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`).join('');

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Rev.io Summit Opportunity Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--navy:#1D3756;--teal:#2399B5;--green:#6EBE4F;--light:#F5F7FA;--border:#DDE2E8;--text:#3d4d5c;--white:#FFFFFF;--ink:#122840;}
*{box-sizing:border-box}body{margin:0;background:linear-gradient(180deg,#10233A 0%,#1D3756 270px,#F5F7FA 271px);color:var(--text);font-family:'Open Sans',Arial,sans-serif}.wrap{max-width:1200px;margin:0 auto;padding:28px 22px 44px}.topbar{height:5px;background:linear-gradient(90deg,var(--teal),var(--green));border-radius:999px;margin-bottom:26px}.hero{display:grid;grid-template-columns:1.35fr .65fr;gap:24px;align-items:end;color:#fff;margin-bottom:24px}.eyebrow{color:#fff;font-weight:800;letter-spacing:.14em;text-transform:uppercase;font-size:12px}.hero h1{font-size:44px;line-height:1.02;margin:10px 0 12px;color:#fff}.hero p{font-size:16px;line-height:1.55;margin:0;max-width:780px;color:#fff}.meta{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.18);border-radius:18px;padding:18px;color:#fff}.meta div{display:flex;justify-content:space-between;gap:14px;border-bottom:1px solid rgba(255,255,255,.18);padding:8px 0}.meta div:last-child{border-bottom:0}.meta b{font-family:Montserrat}.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:14px;margin:20px 0 20px}.card{background:#fff;border:1px solid var(--border);border-radius:18px;padding:18px;box-shadow:0 14px 34px rgba(29,55,86,.08)}.card .label{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#637487;font-weight:800}.card .num{font-family:Montserrat;font-size:30px;font-weight:800;color:var(--navy);margin-top:8px}.card.green .num{color:var(--green)}.card.teal .num{color:var(--teal)}.panel{background:#fff;border:1px solid var(--border);border-radius:20px;padding:20px;margin-top:18px;box-shadow:0 14px 34px rgba(29,55,86,.08)}.panel h2{margin:0 0 14px;color:var(--navy);font-size:22px}.controls{display:grid;grid-template-columns:2fr 1fr 1fr 1fr;gap:10px;margin-bottom:14px}input,select{width:100%;border:1px solid var(--border);border-radius:12px;padding:10px 12px;font:inherit;background:#fff;color:var(--text)}table{width:100%;border-collapse:separate;border-spacing:0}th{text-align:left;color:var(--navy);font-size:12px;text-transform:uppercase;letter-spacing:.08em;border-bottom:2px solid var(--border);padding:12px 10px;white-space:nowrap}td{padding:12px 10px;border-bottom:1px solid var(--border);vertical-align:top}td span{display:block;color:#718095;font-size:12px;margin-top:3px}tbody tr:hover{background:#F8FBFC}.two{display:grid;grid-template-columns:.8fr 1.2fr;gap:18px}.note{font-size:12px;line-height:1.45;color:#637487;margin-top:12px}.hidden{display:none}@media(max-width:900px){.hero,.two{grid-template-columns:1fr}.cards{grid-template-columns:repeat(2,1fr)}.controls{grid-template-columns:1fr}table{font-size:13px}.wrap{padding:18px 12px}.hero h1{font-size:34px}}
</style>
</head>
<body><main class="wrap"><div class="topbar"></div><section class="hero"><div><div class="eyebrow">Rev.io Client Summit 2026</div><h1>Summit Opportunity Dashboard</h1><p>Opportunities created on or after September 1 that are connected to registered Summit attendees through Salesforce Opportunity Contact Roles or the attendee's associated account.</p></div><div class="meta"><div><span>Generated</span><b>${escapeHtml(new Date(summary.generatedAt).toLocaleString('en-US'))}</b></div><div><span>Created since</span><b>${escapeHtml(summary.createdSince.slice(0,10))}</b></div><div><span>Source</span><b>${escapeHtml(summary.sourceFile)}</b></div></div></section>
<section class="cards"><div class="card teal"><div class="label">Unique opps</div><div class="num" id="kpiOpps">${summary.uniqueOpportunities}</div></div><div class="card green"><div class="label">MRR</div><div class="num" id="kpiMrr">${money(summary.totalMRR)}</div></div><div class="card"><div class="label">Accounts</div><div class="num" id="kpiAccounts">${summary.uniqueAccountsWithOpportunities}</div></div><div class="card"><div class="label">Open opps</div><div class="num" id="kpiOpen">${summary.openOpportunities}</div></div><div class="card"><div class="label">Detail rows</div><div class="num" id="kpiRows">${summary.opportunityRows}</div></div></section>
<section class="two"><div class="panel"><h2>Product Type Summary</h2><table><thead><tr><th>Product Type</th><th>Opps</th><th>MRR</th><th>Amount</th></tr></thead><tbody>${productRows}</tbody></table></div><div class="panel"><h2>Filters</h2><div class="controls"><input id="search" placeholder="Search account, attendee, referral, owner…" /><select id="product"><option value="">All product types</option>${options(products)}</select><select id="owner"><option value="">All owners</option>${options(owners)}</select><select id="stage"><option value="">All stages</option>${options(stages)}</select></div><div class="note">MRR uses Opportunity.Total_Monthly_Fees__c first, then Renewal_Amount__c, then Amount as fallback. Rows are one per Salesforce opportunity. If multiple Summit registrants are tied to the same opportunity/account, the dashboard keeps one primary Summit contact, preferring an Opportunity Contact Role match.</div></div></section>
<section class="panel"><h2>Opportunity Detail</h2><table id="detail"><thead><tr><th>Account / Opportunity</th><th>Product type</th><th>MRR</th><th>Contact who attended</th><th>Referred by</th><th>Opp owner</th><th>Stage / Created</th></tr></thead><tbody>${rowHtml || '<tr><td colspan="7">No matching opportunities found.</td></tr>'}</tbody></table><div class="note">Matched ${summary.matchedRegistrants} of ${summary.registrants} active registrants to Salesforce contacts/accounts; ${summary.uniqueMatchedAccounts} unique matched accounts were checked.</div></section>
</main><script>
const rows=[...document.querySelectorAll('#detail tbody tr')];
function parseMoney(s){return Number(String(s).replace(/[^0-9.-]/g,''))||0}
function fmt(n){return n.toLocaleString('en-US',{style:'currency',currency:'USD',maximumFractionDigits:0})}
function apply(){const q=document.getElementById('search').value.trim().toLowerCase();const p=document.getElementById('product').value;const o=document.getElementById('owner').value;const st=document.getElementById('stage').value;let visible=[];rows.forEach(r=>{const show=(!q||(r.dataset.search||'').includes(q))&&(!p||r.dataset.product===p)&&(!o||r.dataset.owner===o)&&(!st||r.dataset.stage===st);r.classList.toggle('hidden',!show);if(show&&r.children.length>1)visible.push(r)});const opps=new Set(),accts=new Set();let mrr=0,open=0;visible.forEach(r=>{const first=r.children[0].innerText;opps.add(first);accts.add(first.split('\n')[0]);mrr+=Number(r.children[2].dataset.num||0);if(!/^closed/i.test(r.dataset.stage||''))open++});document.getElementById('kpiRows').textContent=visible.length;document.getElementById('kpiOpps').textContent=opps.size;document.getElementById('kpiAccounts').textContent=accts.size;document.getElementById('kpiMrr').textContent=fmt(mrr);document.getElementById('kpiOpen').textContent=open}
document.querySelectorAll('input,select').forEach(el=>el.addEventListener('input',apply));
</script></body></html>`;
}

(async () => {
  fs.mkdirSync(dataDir, { recursive: true });
  const registrants = loadRegistrants();
  const token = await getSalesforceToken();
  const enrichedRegistrants = await fetchContactsAndAccounts(token, registrants);
  const opportunities = await fetchOpportunities(token, enrichedRegistrants);
  const rows = buildRows(enrichedRegistrants, opportunities);
  const summary = summarize(rows, enrichedRegistrants);

  fs.writeFileSync(jsonOut, JSON.stringify({ summary, rows }, null, 2));
  writeCsv(rows);
  fs.writeFileSync(htmlOut, renderHtml(summary, rows));
  console.log(JSON.stringify(summary, null, 2));
  console.log(`Wrote ${path.relative(workspaceDir, htmlOut)}`);
  console.log(`Wrote ${path.relative(workspaceDir, csvOut)}`);
  console.log(`Wrote ${path.relative(workspaceDir, jsonOut)}`);
})().catch((err) => { console.error(err); process.exit(1); });
