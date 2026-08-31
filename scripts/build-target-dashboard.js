const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const repoDir = path.resolve(__dirname, '..');
const workspaceDir = path.resolve(repoDir, '..');
const summitDataPath = process.env.SUMMIT_DATA || path.join(repoDir, 'assets', 'data', 'summit-data.json');
const clientCsvPath = process.env.CLIENT_PRODUCTS_CSV || path.join(workspaceDir, 'exports', 'gmail-attachments', '1783436761907-master-client-list-a2654cf0c7e34c57bff58f4ea63d241f_all.csv');
const outDataPath = path.join(repoDir, 'assets', 'data', 'summit-target-data.json');
const outHtmlPath = path.join(repoDir, 'target-dashboard.html');
const outCsvPath = path.join(repoDir, 'assets', 'data', 'summit-target-accounts.csv');

const PRODUCTS = ['Billing', 'PSA Web', 'Tigerpaw', 'Odin', 'Payments'];
const EXCLUDED_CODES = new Set(['REVII', 'SUMMITSPONSOR']);
const COMPANY_SUFFIX_RE = /\b(incorporated|inc|llc|l\.l\.c|ltd|limited|corp|corporation|co|company|communications|communication|telecom|technologies|technology|solutions|services|service|systems|group|direct|usa|c\/o)\b/g;
const GENERIC_SINGLE_MATCH_TOKENS = new Set(['telephone', 'phone', 'voice', 'network', 'networks', 'security', 'secure', 'data', 'digital', 'global', 'premier', 'southeast', 'technology', 'technologies', 'solution', 'solutions', 'system', 'systems']);

function text(value) {
  if (value == null) return '';
  if (typeof value === 'object' && value.text) return String(value.text);
  if (typeof value === 'object' && value.result) return String(value.result);
  return String(value).trim();
}

function normalizeCompany(value) {
  return text(value)
    .toLowerCase()
    .replace(/&/g, ' and ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(COMPANY_SUFFIX_RE, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function canonicalToken(token) {
  let t = token.toLowerCase();
  if (t === 'minutemen') return 'minuteman';
  if (t === 'secured') return 'secure';
  if (t.endsWith('ies') && t.length > 4) return `${t.slice(0, -3)}y`;
  if (t.endsWith('s') && !t.endsWith('ss') && t.length > 4) return t.slice(0, -1);
  return t;
}

function tokens(value) {
  return normalizeCompany(value)
    .split(' ')
    .filter(Boolean)
    .map(canonicalToken)
    .filter(t => t.length > 1 && !['the','and','of'].includes(t));
}

function titleCase(s) {
  return text(s).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).replace(/\b(Ip|It|Msp|Csp|Cfo|Ceo|Coo|Cto|Vp)\b/g, m => m.toUpperCase());
}

function normalizeCode(value) {
  return text(value).toUpperCase().replace(/\s+/g, '');
}

function splitProducts(value) {
  const raw = text(value).toLowerCase();
  const set = new Set();
  if (/billing/.test(raw)) set.add('Billing');
  if (/psa\s*web|psa/.test(raw)) set.add('PSA Web');
  if (/tigerpaw/.test(raw)) set.add('Tigerpaw');
  if (/odin/.test(raw)) set.add('Odin');
  if (/payment/.test(raw)) set.add('Payments');
  return [...set];
}

async function loadClientRows() {
  const wb = new ExcelJS.Workbook();
  await wb.csv.readFile(clientCsvPath);
  const ws = wb.worksheets[0];
  const headers = ws.getRow(1).values.slice(1).map(text);
  const rows = [];
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj = {};
    row.values.slice(1).forEach((value, i) => obj[headers[i]] = text(value));
    if (!obj.Client) return;
    obj.products = splitProducts(obj['Rev.io Product']);
    obj.normalizedClient = normalizeCompany(obj.Client);
    obj.clientTokens = tokens(obj.Client);
    obj.clientCodeTokens = text(obj['Client Code'])
      .split(',')
      .map(part => normalizeCompany(part))
      .filter(Boolean);
    rows.push(obj);
  });
  return rows;
}

function matchClient(company, clientRows) {
  const norm = normalizeCompany(company);
  if (!norm) return null;
  let best = null;
  const companyTokens = new Set(tokens(company));
  for (const client of clientRows) {
    const cn = client.normalizedClient;
    let score = 0;
    const clientTokenList = client.clientTokens || [];
    const clientIsGenericSingle = clientTokenList.length === 1 && GENERIC_SINGLE_MATCH_TOKENS.has(clientTokenList[0]);
    const companyTokenList = [...companyTokens];
    const companyIsGenericSingle = companyTokenList.length === 1 && GENERIC_SINGLE_MATCH_TOKENS.has(companyTokenList[0]);
    if (norm === cn) score = 1;
    else if (client.clientCodeTokens.some(code => code === norm || (code.length >= 4 && norm.includes(code)) || (norm.length >= 4 && code.includes(norm)))) score = 0.96;
    else if (!companyIsGenericSingle && norm.length > 3 && cn.includes(norm)) score = 0.94;
    else if (!clientIsGenericSingle && cn.length > 3 && norm.includes(cn)) score = 0.92;
    else {
      const a = companyTokens;
      const b = new Set(client.clientTokens);
      const intersectionTokens = [...a].filter(t => b.has(t));
      const inter = intersectionTokens.length;
      const union = new Set([...a, ...b]).size || 1;
      const onlyGenericSingle = inter === 1 && GENERIC_SINGLE_MATCH_TOKENS.has(intersectionTokens[0]);
      const jaccard = onlyGenericSingle ? 0 : inter / union;
      const coverage = onlyGenericSingle ? 0 : inter / Math.max(1, Math.min(a.size, b.size));
      score = Math.max(jaccard, coverage * 0.86);
    }
    if (!best || score > best.score) best = { client, score };
  }
  return best && best.score >= 0.55 ? best : null;
}

function htmlEscape(value) {
  return text(value).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function csvEscape(value) {
  const s = text(value);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function main() {
  const summit = JSON.parse(fs.readFileSync(summitDataPath, 'utf8'));
  const clientRows = await loadClientRows();
  const clientCount = clientRows.length;

  const targetedRegistrants = summit.registrants
    .filter(r => normalizeCode(r.discountCode) && !EXCLUDED_CODES.has(normalizeCode(r.discountCode)))
    .map(r => {
      const match = matchClient(r.company, clientRows);
      if (!match) return { ...r, match: null };
      const have = new Set(match.client.products);
      const missing = PRODUCTS.filter(p => !have.has(p));
      return {
        ...r,
        match: {
          client: match.client.Client,
          status: match.client.Status,
          assignedAm: match.client['Assigned AM'],
          averageMrr: match.client['Average MRR'],
          clientCode: match.client['Client Code'],
          products: PRODUCTS.reduce((acc, p) => ({ ...acc, [p]: have.has(p) }), {}),
          productsRaw: match.client['Rev.io Product'],
          missing,
          score: Number(match.score.toFixed(2))
        }
      };
    });

  const accountMap = new Map();
  for (const r of targetedRegistrants.filter(r => r.match)) {
    const key = normalizeCompany(r.match.client) || normalizeCompany(r.company);
    if (!accountMap.has(key)) {
      accountMap.set(key, {
        account: r.match.client,
        registeredCompanyNames: new Set(),
        assignedAm: r.match.assignedAm || 'Unassigned',
        status: r.match.status || '',
        averageMrr: r.match.averageMrr || '',
        clientCode: r.match.clientCode || '',
        products: r.match.products,
        productsRaw: r.match.productsRaw,
        missing: r.match.missing,
        registrants: [],
        referrers: new Set(),
        discountCodes: new Set(),
        matchScore: r.match.score
      });
    }
    const acct = accountMap.get(key);
    acct.registeredCompanyNames.add(r.company);
    acct.registrants.push({
      name: r.name,
      title: r.jobTitle,
      department: r.department,
      rank: r.jobRank,
      email: r.email || '',
      referral: r.referral && !/^none none$/i.test(r.referral) ? r.referral : '',
      discountCode: r.discountCode || '',
      dateRegistered: r.dateLabel || r.dateRegistered || '',
      hotel: r.hotel || '',
      attendedBefore: r.attendedBefore || ''
    });
    if (r.referral && !/^none none$/i.test(r.referral)) acct.referrers.add(r.referral);
    if (r.discountCode) acct.discountCodes.add(r.discountCode);
    acct.matchScore = Math.min(acct.matchScore, r.match.score);
  }

  const accounts = [...accountMap.values()].map(acct => ({
    ...acct,
    registeredCompanyNames: [...acct.registeredCompanyNames].sort(),
    referrers: [...acct.referrers].sort(),
    discountCodes: [...acct.discountCodes].sort(),
    attendeeCount: acct.registrants.length,
    targetOwner: [...acct.referrers].sort().join(', ') || acct.assignedAm || 'Unassigned',
    missingCount: acct.missing.length,
    hasMissingProducts: acct.missing.length > 0
  })).sort((a, b) => (b.missingCount - a.missingCount) || b.attendeeCount - a.attendeeCount || a.account.localeCompare(b.account));

  const referrerRollup = new Map();
  for (const acct of accounts) {
    const owners = acct.referrers.length ? acct.referrers : [acct.assignedAm || 'Unassigned'];
    for (const owner of owners) {
      if (!referrerRollup.has(owner)) referrerRollup.set(owner, { owner, accounts: 0, attendees: 0, missingOpportunities: 0 });
      const row = referrerRollup.get(owner);
      row.accounts += 1;
      row.attendees += acct.attendeeCount;
      row.missingOpportunities += acct.missingCount;
    }
  }
  const rollup = [...referrerRollup.values()].sort((a, b) => b.missingOpportunities - a.missingOpportunities || b.attendees - a.attendees || a.owner.localeCompare(b.owner));

  const unmatched = targetedRegistrants.filter(r => !r.match).map(r => ({
    name: r.name,
    company: r.company,
    referral: r.referral,
    discountCode: r.discountCode,
    dateRegistered: r.dateRegistered || r.dateLabel || ''
  })).sort((a, b) => a.company.localeCompare(b.company));

  const data = {
    generatedAt: new Date().toISOString(),
    source: {
      summitData: path.relative(repoDir, summitDataPath),
      clientProducts: path.relative(repoDir, clientCsvPath),
      liveNotionAccess: false,
      liveNotionNote: 'Notion API returned object_not_found/not shared; used cached CSV export for the same Notion page ID.'
    },
    rules: {
      accountFilter: 'Registrant has a non-empty discount code other than REVII or SUMMITSPONSOR, then company is matched to Master Client List.',
      products: PRODUCTS
    },
    summary: {
      sourceRegistrants: summit.registrants.length,
      discountedNonSponsorRegistrants: targetedRegistrants.length,
      matchedClientRegistrants: targetedRegistrants.filter(r => r.match).length,
      targetAccounts: accounts.length,
      referrers: rollup.length,
      accountsWithMissingProducts: accounts.filter(a => a.hasMissingProducts).length,
      missingProductOpportunities: accounts.reduce((sum, a) => sum + a.missingCount, 0),
      unmatchedRegistrants: unmatched.length,
      clientMasterRows: clientCount
    },
    accounts,
    rollup,
    unmatched
  };
  fs.writeFileSync(outDataPath, JSON.stringify(data, null, 2));

  const csvRows = [
    ['Target Owner','Account','Attendees','Registrant Names','Registered Company Names','Discount Codes','Assigned AM','Current Products','Missing Products',...PRODUCTS,'Average MRR','Client Codes','Match Score']
  ];
  for (const a of accounts) {
    csvRows.push([
      a.targetOwner,
      a.account,
      a.attendeeCount,
      a.registrants.map(r => `${r.name}${r.title ? ` (${r.title})` : ''}`).join('; '),
      a.registeredCompanyNames.join('; '),
      a.discountCodes.join('; '),
      a.assignedAm,
      PRODUCTS.filter(p => a.products[p]).join(', '),
      a.missing.join(', '),
      ...PRODUCTS.map(p => a.products[p] ? 'Yes' : 'No'),
      a.averageMrr,
      a.clientCode,
      a.matchScore
    ]);
  }
  fs.writeFileSync(outCsvPath, csvRows.map(row => row.map(csvEscape).join(',')).join('\n'));

  const html = buildHtml(data);
  fs.writeFileSync(outHtmlPath, html);
  console.log(`Built ${path.relative(repoDir, outHtmlPath)}`);
  console.log(`Built ${path.relative(repoDir, outDataPath)}`);
  console.log(`Built ${path.relative(repoDir, outCsvPath)}`);
  console.log(JSON.stringify(data.summary, null, 2));
}

function buildHtml(data) {
  const stat = (label, value, sub='') => `<div class="stat"><div class="stat-value">${htmlEscape(value)}</div><div class="stat-label">${htmlEscape(label)}</div>${sub ? `<div class="stat-sub">${htmlEscape(sub)}</div>` : ''}</div>`;
  const productHeader = PRODUCTS.map(p => `<th>${htmlEscape(p)}</th>`).join('');
  const productCells = a => PRODUCTS.map(p => `<td class="product ${a.products[p] ? 'yes' : 'no'}">${a.products[p] ? '✓' : '—'}</td>`).join('');
  const rows = data.accounts.map(a => `
    <tr data-owner="${htmlEscape(a.targetOwner.toLowerCase())}" data-products="${htmlEscape(PRODUCTS.filter(p => a.products[p]).join(' ').toLowerCase())}" data-missing="${htmlEscape(a.missing.join(' ').toLowerCase())}" data-search="${htmlEscape([a.account, a.targetOwner, a.assignedAm, a.registeredCompanyNames.join(' '), a.registrants.map(r => r.name).join(' '), a.discountCodes.join(' ')].join(' ').toLowerCase())}">
      <td class="sticky"><strong>${htmlEscape(a.account)}</strong><span>${htmlEscape(a.registeredCompanyNames.join(' / '))}</span></td>
      <td>${htmlEscape(a.targetOwner || 'Unassigned')}<span>Assigned AM: ${htmlEscape(a.assignedAm || '—')}</span></td>
      <td class="num">${a.attendeeCount}<span>${htmlEscape(a.registrants.map(r => r.name).join('; '))}</span></td>
      <td>${htmlEscape(a.discountCodes.join(', ') || '—')}</td>
      ${productCells(a)}
      <td class="missing">${a.missing.length ? a.missing.map(p => `<b>${htmlEscape(p)}</b>`).join(' ') : '<em>Complete set</em>'}</td>
      <td>${htmlEscape(a.averageMrr || '—')}</td>
      <td><button class="details" type="button">View</button></td>
    </tr>
    <tr class="detail-row"><td colspan="${10 + PRODUCTS.length}"><div class="details-box">
      <div><strong>Registrant targeting notes</strong><ul>${a.registrants.map(r => `<li>${htmlEscape(r.name)}${r.title ? ` — ${htmlEscape(r.title)}` : ''}${r.department ? ` · ${htmlEscape(r.department)}` : ''}${r.referral ? ` · referred by ${htmlEscape(r.referral)}` : ''}${r.attendedBefore ? ` · attended before: ${htmlEscape(r.attendedBefore)}` : ''}</li>`).join('')}</ul></div>
      <div><strong>Client codes</strong><p>${htmlEscape(a.clientCode || '—')}</p><strong>Raw product field</strong><p>${htmlEscape(a.productsRaw || '—')}</p><strong>Match score</strong><p>${htmlEscape(a.matchScore)}</p></div>
    </div></td></tr>`).join('');

  const rollupRows = data.rollup.map(r => `<tr><td>${htmlEscape(r.owner)}</td><td class="num">${r.accounts}</td><td class="num">${r.attendees}</td><td class="num">${r.missingOpportunities}</td></tr>`).join('');

  const unmatchedRows = data.unmatched.map(u => `<tr><td>${htmlEscape(u.company)}</td><td>${htmlEscape(u.name)}</td><td>${htmlEscape(u.referral || '—')}</td><td>${htmlEscape(u.discountCode || '—')}</td></tr>`).join('');

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Rev.io Summit Client Target Dashboard</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--navy:#1D3756;--teal:#2399B5;--green:#6EBE4F;--light:#F5F7FA;--border:#DDE2E8;--body:#3d4d5c;--white:#FFFFFF;}
*{box-sizing:border-box} body{margin:0;background:var(--light);color:var(--body);font-family:'Open Sans',Arial,sans-serif;} .topbar{height:5px;background:linear-gradient(90deg,var(--teal),var(--green));}
.hero{background:radial-gradient(circle at 15% 20%,rgba(35,153,181,.34),transparent 34%),radial-gradient(circle at 82% 5%,rgba(110,190,79,.22),transparent 30%),linear-gradient(135deg,#112a43,var(--navy));color:var(--white);padding:34px 28px 46px;}
.wrap{max-width:1400px;margin:0 auto}.eyebrow{font:800 12px/1 'Montserrat';letter-spacing:.18em;text-transform:uppercase;color:var(--green);}.hero h1{margin:12px 0 8px;font-size:clamp(30px,4vw,56px);line-height:1.02;color:var(--white);}.hero p{max-width:920px;margin:0;color:var(--white);font-size:17px;line-height:1.5}.meta{margin-top:18px;display:flex;flex-wrap:wrap;gap:10px}.pill{border:1px solid rgba(255,255,255,.28);border-radius:999px;padding:7px 12px;color:var(--white);font-size:12px;font-weight:700;background:rgba(255,255,255,.08)}
main{max-width:1400px;margin:-26px auto 60px;padding:0 20px}.stats{display:grid;grid-template-columns:repeat(6,minmax(140px,1fr));gap:14px}.stat{background:var(--white);border:1px solid var(--border);border-radius:18px;padding:18px;box-shadow:0 10px 30px rgba(29,55,86,.08)}.stat-value{font-family:'Montserrat';font-weight:800;font-size:30px;color:var(--navy)}.stat-label{font-weight:800;font-size:12px;text-transform:uppercase;letter-spacing:.06em}.stat-sub{font-size:12px;margin-top:4px;color:#647386}
.panel{background:var(--white);border:1px solid var(--border);border-radius:20px;margin-top:18px;padding:18px;box-shadow:0 10px 30px rgba(29,55,86,.06)}.panel h2{color:var(--navy);margin:0 0 12px;font-size:21px}.controls{display:grid;grid-template-columns:2fr repeat(3,1fr);gap:10px;margin-bottom:14px}.controls input,.controls select{border:1px solid var(--border);border-radius:12px;padding:11px 12px;font:inherit;color:var(--body);background:white}.actions{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap}.download{display:inline-block;background:var(--green);color:#113024;text-decoration:none;font-weight:800;border-radius:12px;padding:10px 14px}.note{font-size:12px;color:#647386}.table-wrap{overflow:auto;border:1px solid var(--border);border-radius:16px}table{width:100%;border-collapse:separate;border-spacing:0;min-width:1280px}th{position:sticky;top:0;background:#eef6f8;color:var(--navy);font-size:12px;text-align:left;text-transform:uppercase;letter-spacing:.04em;padding:12px;border-bottom:1px solid var(--border);z-index:2}td{padding:12px;border-bottom:1px solid var(--border);vertical-align:top;font-size:13px}td span{display:block;color:#6b7787;font-size:11px;margin-top:3px;max-width:320px}.sticky{position:sticky;left:0;background:white;z-index:1;box-shadow:1px 0 0 var(--border)}.num{text-align:right;font-family:'Montserrat';font-weight:700;color:var(--navy)}.product{text-align:center;font:800 18px/1 'Montserrat'}.product.yes{color:var(--green)}.product.no{color:#b6c0cb}.missing b{display:inline-block;margin:0 4px 4px 0;background:#eaf6ea;color:#2d6530;border:1px solid #cce8c8;border-radius:999px;padding:4px 8px;font-size:11px}.missing em{color:#7a8795}.details{border:0;background:var(--teal);color:white;border-radius:10px;padding:7px 11px;font-weight:800;cursor:pointer}.detail-row{display:none}.detail-row.open{display:table-row}.details-box{display:grid;grid-template-columns:2fr 1fr;gap:20px;background:#f8fbfc;border-radius:14px;padding:14px}.details-box ul{margin:8px 0 0;padding-left:18px}.details-box p{margin:6px 0 12px}.rollup{max-width:680px;min-width:520px}.warning{border-left:5px solid var(--teal);background:#f6fbfc}.footer{margin:22px 0;color:#6b7787;font-size:12px}@media(max-width:900px){.stats{grid-template-columns:repeat(2,1fr)}.controls{grid-template-columns:1fr}.details-box{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="topbar"></div>
<header class="hero"><div class="wrap"><div class="eyebrow">Rev.io Summit 2026 · Client Targeting</div><h1>Client target dashboard</h1><p>Account-level view of Summit attendees who used non-sponsor/non-REVII discount codes, matched to the Master Client List products so referrers know which clients they own and which products to target onsite.</p><div class="meta"><span class="pill">Excludes REVII</span><span class="pill">Excludes SUMMITSPONSOR</span><span class="pill">Products: Billing · PSA Web · Tigerpaw · Odin · Payments</span><span class="pill">Generated ${htmlEscape(new Date(data.generatedAt).toLocaleString('en-US', { timeZone: 'UTC', dateStyle: 'medium', timeStyle: 'short' }))} UTC</span></div></div></header>
<main>
<section class="stats">
${stat('Target accounts', data.summary.targetAccounts)}${stat('Matched attendees', data.summary.matchedClientRegistrants, `${data.summary.discountedNonSponsorRegistrants} discounted non-sponsor registrants`)}${stat('Referrers / owners', data.summary.referrers)}${stat('Accounts missing products', data.summary.accountsWithMissingProducts)}${stat('Missing product opps', data.summary.missingProductOpportunities)}${stat('Unmatched registrants', data.summary.unmatchedRegistrants, 'Need manual account match')}
</section>
<section class="panel"><div class="actions"><div><h2>Target account list</h2><div class="note">Owner = registrant referral name when present; falls back to Assigned AM. Product checks come from Master Client List “Rev.io Product”.</div></div><a class="download" href="assets/data/summit-target-accounts.csv">Download CSV</a></div>
<div class="controls"><input id="search" placeholder="Search account, attendee, owner, code…"><select id="owner"><option value="">All owners</option>${data.rollup.map(r => `<option>${htmlEscape(r.owner)}</option>`).join('')}</select><select id="missing"><option value="">All missing products</option>${PRODUCTS.map(p => `<option>${htmlEscape(p)}</option>`).join('')}<option value="none">No missing products</option></select><select id="have"><option value="">All current products</option>${PRODUCTS.map(p => `<option>${htmlEscape(p)}</option>`).join('')}</select></div>
<div class="table-wrap"><table id="accounts"><thead><tr><th class="sticky">Account</th><th>Target owner</th><th>Attendees</th><th>Codes</th>${productHeader}<th>Missing / target</th><th>Avg MRR</th><th>Details</th></tr></thead><tbody>${rows}</tbody></table></div></section>
<section class="panel"><h2>Referrer workload</h2><div class="table-wrap rollup"><table><thead><tr><th>Owner</th><th>Accounts</th><th>Attendees</th><th>Missing product opps</th></tr></thead><tbody>${rollupRows}</tbody></table></div></section>
<section class="panel warning"><h2>Unmatched discounted registrants</h2><p class="note">These passed the discount-code rule but did not confidently match a Master Client List account. They are excluded from target account stats until manually mapped.</p><div class="table-wrap"><table><thead><tr><th>Company</th><th>Registrant</th><th>Referral</th><th>Code</th></tr></thead><tbody>${unmatchedRows || '<tr><td colspan="4">No unmatched registrants.</td></tr>'}</tbody></table></div></section>
<div class="footer">Source note: Live Notion page access returned not shared/object_not_found, so this build used the cached Master Client List CSV export with matching page ID.</div>
</main>
<script>
const rows=[...document.querySelectorAll('#accounts tbody tr:not(.detail-row)')];
function filter(){const q=document.querySelector('#search').value.toLowerCase().trim();const owner=document.querySelector('#owner').value.toLowerCase();const missing=document.querySelector('#missing').value.toLowerCase();const have=document.querySelector('#have').value.toLowerCase();for(const tr of rows){const detail=tr.nextElementSibling;let ok=true;if(q&&!tr.dataset.search.includes(q))ok=false;if(owner&&!tr.dataset.owner.includes(owner))ok=false;if(missing==='none'&&tr.dataset.missing)ok=false;else if(missing&&missing!=='none'&&!tr.dataset.missing.includes(missing))ok=false;if(have&&!tr.dataset.products.includes(have))ok=false;tr.style.display=ok?'':'none';if(!ok)detail.classList.remove('open');}}
document.querySelectorAll('.controls input,.controls select').forEach(el=>el.addEventListener('input',filter));
document.querySelectorAll('.details').forEach(btn=>btn.addEventListener('click',e=>e.target.closest('tr').nextElementSibling.classList.toggle('open')));
</script>
</body></html>`;
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
