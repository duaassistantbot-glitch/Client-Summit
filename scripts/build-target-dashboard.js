const fs = require('fs');
const path = require('path');
const https = require('https');
const ExcelJS = require('exceljs');

const repoDir = path.resolve(__dirname, '..');
const workspaceDir = path.resolve(repoDir, '..');
const summitDataPath = process.env.SUMMIT_DATA || path.join(repoDir, 'assets', 'data', 'summit-data.json');
const clientCsvPath = process.env.CLIENT_PRODUCTS_CSV || path.join(workspaceDir, 'exports', 'gmail-attachments', '1783436761907-master-client-list-a2654cf0c7e34c57bff58f4ea63d241f_all.csv');
const outDataPath = path.join(repoDir, 'assets', 'data', 'summit-target-data.json');
const outHtmlPath = path.join(repoDir, 'target-dashboard.html');
const outCsvPath = path.join(repoDir, 'assets', 'data', 'summit-target-accounts.csv');

const PRODUCTS = ['Billing', 'New Rev.io', 'Payments'];
const PRODUCT_SOURCE_LABELS = { 'New Rev.io': 'PSA Web' };
const API_VERSION = 'v59.0';
const EXCLUDED_CODES = new Set(['REVII', 'SUMMITSPONSOR']);
const EXCLUDED_NON_TARGET_COMPANIES = new Set(['ooma', 'kealywalker', 'yorn sales training', 'crancer', 'the crancer']);
const MANUAL_SF_LOOKUP_NAMES = new Map([
  ['aimerica', 'Empire Telecom'],
  ['southeast telephone', 'SouthEast Telephone'],
  ['true choice', 'Blueline Telecom']
]);
const COMPANY_SUFFIX_RE = /\b(incorporated|inc|llc|l\.l\.c|ltd|limited|corp|corporation|co|company|communications|communication|telecom|technologies|technology|solutions|services|service|systems|group|direct|usa|c\/o)\b/g;
const GENERIC_SINGLE_MATCH_TOKENS = new Set(['telephone', 'phone', 'voice', 'network', 'networks', 'security', 'secure', 'data', 'digital', 'global', 'premier', 'southeast', 'technology', 'technologies', 'solution', 'solutions', 'system', 'systems']);
const GENERIC_CLIENT_CODE_TOKENS = new Set(['demo', 'training', 'sales', 'template', 'product', 'solutions', 'solution', 'inventory', 'sandbox', 'testdrive', 'learn']);

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

function manualSfLookupName(company) {
  const norm = normalizeCompany(company);
  for (const [key, value] of MANUAL_SF_LOOKUP_NAMES.entries()) {
    if (norm === key || norm.includes(key)) return value;
  }
  return text(company);
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
  if (/psa\s*web|psa/.test(raw)) set.add('New Rev.io');
  if (/payment/.test(raw)) set.add('Payments');
  return [...set];
}

function productsFromSfSignals(sfAccount, opps = []) {
  const set = new Set();
  const fields = [sfAccount?.Billing_Platform__c, sfAccount?.Current_Platform__c, sfAccount?.PSA_Platform__c, ...(opps.filter(o => o.IsWon).flatMap(o => [o.Name, o.Type, o.Product_Type__c]))].join(' ').toLowerCase();
  if (/billing|rev\.io|revio/.test(fields)) set.add('Billing');
  if (/psa|new rev\.io|new revio/.test(fields) || sfAccount?.PSA_Web__c) set.add('New Rev.io');
  if (/payment/.test(fields)) set.add('Payments');
  return [...set];
}

function loadEnv() {
  const envPath = path.join(workspaceDir, '.env');
  if (!fs.existsSync(envPath)) return;
  for (const line of fs.readFileSync(envPath, 'utf8').split('\n')) {
    const match = line.match(/^([^#=]+)=(.*)$/);
    if (match && !process.env[match[1].trim()]) process.env[match[1].trim()] = match[2].trim();
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
      response.on('data', chunk => data += chunk);
      response.on('end', () => {
        let parsed = data;
        try { parsed = data ? JSON.parse(data) : null; } catch {}
        if (response.statusCode >= 400) {
          const message = Array.isArray(parsed) ? parsed.map(item => item.message || item.errorCode).join('; ') : parsed?.message || parsed?.error_description || data || `HTTP ${response.statusCode}`;
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
  if (!process.env.SF_CLIENT_ID || !process.env.SF_CLIENT_SECRET || !process.env.SF_INSTANCE_URL) return null;
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

function soqlString(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

async function sfQuery(token, soql) {
  const records = [];
  let requestPath = `/services/data/${API_VERSION}/query?q=${encodeURIComponent(soql)}`;
  while (requestPath) {
    const result = await sfRequest({ requestPath, token });
    records.push(...(result.records || []));
    requestPath = result.nextRecordsUrl || null;
  }
  return records;
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
    else if (client.clientCodeTokens.some(code => {
      const codeIsGeneric = GENERIC_CLIENT_CODE_TOKENS.has(code);
      return code === norm || (!codeIsGeneric && ((code.length >= 4 && norm.includes(code)) || (norm.length >= 4 && code.includes(norm))));
    })) score = 0.96;
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

async function loadSalesforceAccountData(accounts) {
  let token;
  try {
    token = await getSalesforceToken();
  } catch (error) {
    console.warn(`Salesforce auth failed; cohort badges unavailable: ${error.message}`);
    return { byKey: new Map(), available: false, error: error.message };
  }
  if (!token) return { byKey: new Map(), available: false, error: 'Salesforce credentials unavailable' };

  const directIds = [...new Set(accounts.map(account => account.sfAccountId).filter(Boolean))];
  const directIdClause = directIds.length ? `Id IN (${directIds.map(id => `'${soqlString(id)}'`).join(',')}) OR` : '';
  const sfAccounts = await sfQuery(token, `
    SELECT Id, Name, Client_Code__c, Owner.Name, TigerPaw_Account_Status__c, Tigerpaw__c, Odin__c,
           Odin_Account_Status__c, Broadsoft_Type__c, Billing_Platform__c, Current_Platform__c,
           PSA_Web__c, PSA_Platform__c
    FROM Account
    WHERE ${directIdClause} Client_Code__c != null
       OR TigerPaw_Account_Status__c != null
       OR Tigerpaw__c = true
       OR Odin__c = true
       OR Odin_Account_Status__c != null
       OR PSA_Web__c = true
       OR Broadsoft_Type__c != null
  `);

  const matched = new Map();
  const accountIds = new Set();
  for (const account of accounts) {
    const sfMatch = matchSalesforceAccount(account, sfAccounts);
    if (sfMatch) {
      matched.set(account.account, sfMatch);
      accountIds.add(sfMatch.Id);
    }
  }

  const oppsByAccount = new Map();
  const ids = [...accountIds];
  for (let i = 0; i < ids.length; i += 80) {
    const chunk = ids.slice(i, i + 80).map(id => `'${soqlString(id)}'`).join(',');
    const opps = await sfQuery(token, `
      SELECT Id, AccountId, Name, Type, Product_Type__c, StageName, IsWon, CloseDate, CreatedDate
      FROM Opportunity
      WHERE AccountId IN (${chunk}) AND IsWon = true
      ORDER BY CloseDate ASC, CreatedDate ASC
    `);
    for (const opp of opps) {
      if (!oppsByAccount.has(opp.AccountId)) oppsByAccount.set(opp.AccountId, []);
      oppsByAccount.get(opp.AccountId).push(opp);
    }
  }

  const eventsByAccount = new Map();
  for (let i = 0; i < ids.length; i += 80) {
    const chunk = ids.slice(i, i + 80).map(id => `'${soqlString(id)}'`).join(',');
    const events = await sfQuery(token, `
      SELECT Id, AccountId, Subject, StartDateTime, EndDateTime, ActivityDate, Owner.Name
      FROM Event
      WHERE AccountId IN (${chunk})
        AND ActivityDate >= 2026-09-01
        AND ActivityDate <= 2026-09-03
      ORDER BY StartDateTime ASC
    `);
    for (const event of events) {
      if (!eventsByAccount.has(event.AccountId)) eventsByAccount.set(event.AccountId, []);
      eventsByAccount.get(event.AccountId).push({
        subject: event.Subject || 'Meeting',
        date: event.ActivityDate || '',
        startDateTime: event.StartDateTime || '',
        endDateTime: event.EndDateTime || '',
        owner: event.Owner?.Name || ''
      });
    }
  }

  const byKey = new Map();
  for (const [accountName, sfAccount] of matched.entries()) {
    byKey.set(accountName, {
      sfAccountId: sfAccount.Id,
      sfAccountName: sfAccount.Name,
      accountOwner: sfAccount.Owner?.Name || '',
      cohort: determineCohort(sfAccount, oppsByAccount.get(sfAccount.Id) || []),
      usesBroadworks: hasBroadworks(sfAccount, oppsByAccount.get(sfAccount.Id) || []),
      summitEvents: eventsByAccount.get(sfAccount.Id) || []
    });
  }
  return { byKey, available: true, matchedAccounts: byKey.size, queriedAccounts: sfAccounts.length, eventAccounts: eventsByAccount.size, eventCount: [...eventsByAccount.values()].reduce((sum, rows) => sum + rows.length, 0) };
}

async function loadSalesforceFallbackClientRows(companies, existingClientRows) {
  const unmatchedCompanies = [...new Set(companies)]
    .filter(company => !matchClient(company, existingClientRows))
    .filter(company => !EXCLUDED_NON_TARGET_COMPANIES.has(normalizeCompany(company)));
  if (!unmatchedCompanies.length) return { rows: [], source: { attemptedCompanies: 0, matchedCompanies: 0 } };

  let token;
  try {
    token = await getSalesforceToken();
  } catch (error) {
    console.warn(`Salesforce fallback matching failed: ${error.message}`);
    return { rows: [], source: { attemptedCompanies: unmatchedCompanies.length, matchedCompanies: 0, error: error.message } };
  }
  if (!token) return { rows: [], source: { attemptedCompanies: unmatchedCompanies.length, matchedCompanies: 0, error: 'Salesforce credentials unavailable' } };

  const rows = [];
  for (const company of unmatchedCompanies) {
    const lookupName = manualSfLookupName(company);
    const candidates = await findSalesforceAccountCandidates(token, company, lookupName);
    const best = chooseBestSalesforceCandidate(company, lookupName, candidates);
    if (!best) continue;

    const scoredCandidates = [];
    for (const candidate of candidates.filter(c => normalizeCompany(c.Name) === normalizeCompany(best.Name))) {
      const candidateOpps = await sfQuery(token, `
        SELECT Id, AccountId, Name, Type, Product_Type__c, StageName, IsWon, CloseDate, CreatedDate
        FROM Opportunity
        WHERE AccountId = '${soqlString(candidate.Id)}'
        ORDER BY CloseDate ASC, CreatedDate ASC
      `);
      const candidateProducts = productsFromSfSignals(candidate, candidateOpps);
      scoredCandidates.push({ candidate, opps: candidateOpps, products: candidateProducts });
    }
    const selected = scoredCandidates.sort((a, b) => b.products.length - a.products.length || b.opps.filter(o => o.IsWon).length - a.opps.filter(o => o.IsWon).length)[0] || { candidate: best, opps: [], products: [] };
    const bestAccount = selected.candidate;
    const products = selected.products;
    rows.push({
      Client: company,
      Status: '',
      'Assigned AM': bestAccount.Owner?.Name || 'Unassigned',
      'Rev.io Product': products.join(', '),
      products,
      normalizedClient: normalizeCompany(company),
      clientTokens: tokens(company),
      clientCodeTokens: [],
      sfFallback: true,
      sfLookupName: lookupName,
      sfAccountId: bestAccount.Id,
      sfAccountName: bestAccount.Name
    });
  }
  return { rows, source: { attemptedCompanies: unmatchedCompanies.length, matchedCompanies: rows.length } };
}

async function findSalesforceAccountCandidates(token, company, lookupName) {
  const terms = [...new Set([lookupName, company, text(company).replace(/, Inc\.?$/i, ''), text(company).split(/\s+/)[0]].filter(Boolean))];
  const byId = new Map();
  for (const term of terms) {
    const like = `%${soqlString(term)}%`;
    const accounts = await sfQuery(token, `
      SELECT Id, Name, Owner.Name, Client_Code__c, TigerPaw_Account_Status__c, Tigerpaw__c, Odin__c,
             Odin_Account_Status__c, Broadsoft_Type__c, Billing_Platform__c, Current_Platform__c,
             PSA_Web__c, PSA_Platform__c
      FROM Account
      WHERE Name LIKE '${like}' OR Client_Code__c LIKE '${like}'
      LIMIT 20
    `);
    for (const account of accounts) byId.set(account.Id, account);

    const opps = await sfQuery(token, `
      SELECT AccountId, Account.Name, Account.Owner.Name, Account.Client_Code__c, Account.TigerPaw_Account_Status__c,
             Account.Tigerpaw__c, Account.Odin__c, Account.Odin_Account_Status__c, Account.Broadsoft_Type__c,
             Account.Billing_Platform__c, Account.Current_Platform__c, Account.PSA_Web__c, Account.PSA_Platform__c
      FROM Opportunity
      WHERE Name LIKE '${like}' OR Client_Name__c LIKE '${like}' OR Account.Name LIKE '${like}'
      ORDER BY IsWon DESC, CloseDate DESC
      LIMIT 20
    `);
    for (const opp of opps) {
      if (!opp.AccountId || !opp.Account) continue;
      byId.set(opp.AccountId, { Id: opp.AccountId, ...opp.Account });
    }
  }
  return [...byId.values()];
}

function chooseBestSalesforceCandidate(company, lookupName, candidates) {
  if (!candidates.length) return null;
  const companyNorm = normalizeCompany(company);
  const lookupNorm = normalizeCompany(lookupName);
  let best = null;
  for (const account of candidates) {
    const nameNorm = normalizeCompany(account.Name);
    let score = 0;
    if (nameNorm === lookupNorm) score = 1;
    else if (nameNorm === companyNorm) score = 0.99;
    else if (lookupNorm.length > 3 && nameNorm.includes(lookupNorm)) score = 0.92;
    else if (companyNorm.length > 3 && nameNorm.includes(companyNorm)) score = 0.9;
    else {
      const a = new Set(tokens(lookupName));
      const b = new Set(tokens(account.Name));
      const inter = [...a].filter(t => b.has(t)).length;
      score = inter / Math.max(1, Math.min(a.size, b.size));
    }
    if (!best || score > best.score) best = { ...account, score };
  }
  return best && best.score >= 0.75 ? best : null;
}

function matchSalesforceAccount(account, sfAccounts) {
  if (account.sfAccountId) return sfAccounts.find(sf => sf.Id === account.sfAccountId) || null;
  const codeParts = text(account.clientCode).split(',').map(c => normalizeCompany(c)).filter(Boolean);
  let best = null;
  for (const sf of sfAccounts) {
    const sfCodes = text(sf.Client_Code__c).split(',').map(c => normalizeCompany(c)).filter(Boolean);
    let score = 0;
    if (codeParts.length && sfCodes.length && codeParts.some(c => sfCodes.includes(c))) score = 1;
    else {
      const match = matchClient(sf.Name, [{
        Client: account.account,
        normalizedClient: normalizeCompany(account.account),
        clientTokens: tokens(account.account),
        clientCodeTokens: codeParts
      }]);
      score = match?.score || 0;
    }
    if (!best || score > best.score) best = { ...sf, score };
  }
  return best && best.score >= 0.55 ? best : null;
}

function hasBroadworks(sfAccount, opps) {
  const haystack = [sfAccount.Broadsoft_Type__c, sfAccount.Current_Platform__c, sfAccount.Billing_Platform__c, sfAccount.PSA_Platform__c, ...opps.flatMap(o => [o.Name, o.Type, o.Product_Type__c])].join(' ').toLowerCase();
  return /broadworks|broadsoft/.test(haystack);
}

function inferFallbackCohort(productsRaw) {
  const raw = text(productsRaw).toLowerCase();
  if (/tigerpaw/.test(raw)) return 'Tigerpaw';
  if (/odin|broadworks|broadsoft/.test(raw)) return 'Odin';
  return 'Rev.io Billing';
}

function determineCohort(sfAccount, opps) {
  if (text(sfAccount.TigerPaw_Account_Status__c)) return 'Tigerpaw';
  const firstWon = [...opps].sort((a, b) => text(a.CloseDate).localeCompare(text(b.CloseDate)) || text(a.CreatedDate).localeCompare(text(b.CreatedDate)))[0];
  const firstText = firstWon ? [firstWon.Name, firstWon.Type, firstWon.Product_Type__c].join(' ').toLowerCase() : '';
  if (/tigerpaw|psa/.test(firstText)) return 'Tigerpaw';
  if (/odin|broadworks|broadsoft/.test(firstText) || sfAccount.Odin__c || text(sfAccount.Odin_Account_Status__c)) return 'Odin';
  if (/billing|rev\.io|revio/.test(firstText)) return 'Rev.io Billing';
  if (hasBroadworks(sfAccount, opps)) return 'Odin';
  return 'Rev.io Billing';
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
  const eligibleRegistrants = summit.registrants
    .filter(r => normalizeCode(r.discountCode) && !EXCLUDED_CODES.has(normalizeCode(r.discountCode)))
    .filter(r => !EXCLUDED_NON_TARGET_COMPANIES.has(normalizeCompany(r.company)));
  const sfFallback = await loadSalesforceFallbackClientRows(eligibleRegistrants.map(r => r.company), clientRows);
  const allClientRows = [...clientRows, ...sfFallback.rows];

  const targetedRegistrants = eligibleRegistrants
    .map(r => {
      const match = matchClient(r.company, allClientRows);
      if (!match) return { ...r, match: null };
      const have = new Set(match.client.products);
      const missing = PRODUCTS.filter(p => !have.has(p));
      return {
        ...r,
        match: {
          client: match.client.Client,
          status: match.client.Status,
          assignedAm: match.client['Assigned AM'],
          sfFallback: Boolean(match.client.sfFallback),
          sfLookupName: match.client.sfLookupName || '',
          sfAccountId: match.client.sfAccountId || '',
          sfAccountName: match.client.sfAccountName || '',
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
        clientCode: r.match.clientCode || '',
        sfFallback: Boolean(r.match.sfFallback),
        sfLookupName: r.match.sfLookupName || '',
        sfAccountId: r.match.sfAccountId || '',
        sfAccountName: r.match.sfAccountName || '',
        products: r.match.products,
        productsRaw: r.match.productsRaw,
        missing: r.match.missing,
        registrants: [],
        referrers: new Set(),
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
      dateRegistered: r.dateLabel || r.dateRegistered || '',
      hotel: r.hotel || '',
      attendedBefore: r.attendedBefore || ''
    });
    if (r.referral && !/^none none$/i.test(r.referral)) acct.referrers.add(r.referral);
    acct.matchScore = Math.min(acct.matchScore, r.match.score);
  }

  const accounts = [...accountMap.values()].map(acct => ({
    ...acct,
    registeredCompanyNames: [...acct.registeredCompanyNames].sort(),
    referrers: [...acct.referrers].sort(),
    attendeeCount: acct.registrants.length,
    referralOwner: [...acct.referrers].sort().join(', ') || 'Unassigned',
    targetOwner: [...acct.referrers].sort().join(', ') || acct.assignedAm || 'Unassigned',
    originalCohort: 'Checking Salesforce…',
    accountOwner: 'Checking Salesforce…',
    usesBroadworks: false,
    summitEvents: [],
    missingCount: acct.missing.length,
    hasMissingProducts: acct.missing.length > 0
  })).sort((a, b) => (b.missingCount - a.missingCount) || b.attendeeCount - a.attendeeCount || a.account.localeCompare(b.account));

  const sfCohorts = await loadSalesforceAccountData(accounts);
  for (const account of accounts) {
    const sf = sfCohorts.byKey.get(account.account);
    account.originalCohort = sf?.cohort || inferFallbackCohort(account.productsRaw);
    account.accountOwner = sf?.accountOwner || 'Unassigned';
    account.usesBroadworks = Boolean(sf?.usesBroadworks);
    account.summitEvents = sf?.summitEvents || [];
    account.sfAccountId = sf?.sfAccountId || '';
    account.sfAccountName = sf?.sfAccountName || '';
    delete account.clientCode;
  }

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

  const ignoredSponsorRegistrants = summit.registrants
    .filter(r => normalizeCode(r.discountCode) && !EXCLUDED_CODES.has(normalizeCode(r.discountCode)))
    .filter(r => EXCLUDED_NON_TARGET_COMPANIES.has(normalizeCompany(r.company)))
    .map(r => ({
      name: r.name,
      company: r.company,
      referral: r.referral,
      dateRegistered: r.dateRegistered || r.dateLabel || ''
    }))
    .sort((a, b) => a.company.localeCompare(b.company));

  const unmatched = targetedRegistrants
    .filter(r => !r.match && !EXCLUDED_NON_TARGET_COMPANIES.has(normalizeCompany(r.company)))
    .map(r => ({
      name: r.name,
      company: r.company,
      referral: r.referral,
      dateRegistered: r.dateRegistered || r.dateLabel || ''
    })).sort((a, b) => a.company.localeCompare(b.company));

  const data = {
    generatedAt: new Date().toISOString(),
    source: {
      summitData: path.relative(repoDir, summitDataPath),
      clientProducts: path.relative(repoDir, clientCsvPath),
      salesforceCohorts: sfCohorts,
      salesforceFallbackMatches: sfFallback.source,
      liveNotionAccess: false,
      liveNotionNote: 'Notion API returned object_not_found/not shared; used cached CSV export for the same Notion page ID.'
    },
    rules: {
      accountFilter: 'Registrant has a non-empty discount code other than REVII or SUMMITSPONSOR, then company is matched to Master Client List with Salesforce Account/Opportunity fallback. Ooma, KealyWalker, YorN Sales Training, and The Crancer Group are ignored as non-target sponsor/speaker records.',
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
      ignoredSponsorRegistrants: ignoredSponsorRegistrants.length,
      sfFallbackMatchedAccounts: sfFallback.rows.length,
      clientMasterRows: clientCount
    },
    accounts,
    rollup,
    unmatched,
    ignoredSponsorRegistrants
  };
  fs.writeFileSync(outDataPath, JSON.stringify(data, null, 2));

  const csvRows = [
    ['Referral Owner','Account Owner','AM Owner','Account','Original Cohort','Summit Events Sep 1-3','Attendees','Registrant Names','Registered Company Names','Current Products','Missing Products',...PRODUCTS,'Match Score']
  ];
  for (const a of accounts) {
    csvRows.push([
      a.referralOwner,
      a.accountOwner,
      a.assignedAm,
      a.account,
      a.originalCohort,
      a.summitEvents.map(e => `${e.date} ${e.subject}${e.owner ? ` (${e.owner})` : ''}`).join('; '),
      a.attendeeCount,
      a.registrants.map(r => `${r.name}${r.title ? ` (${r.title})` : ''}`).join('; '),
      a.registeredCompanyNames.join('; '),
      PRODUCTS.filter(p => a.products[p]).join(', '),
      a.missing.join(', '),
      ...PRODUCTS.map(p => a.products[p] ? 'Yes' : 'No'),
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
  const productCells = a => PRODUCTS.map(p => `<td data-label="${htmlEscape(p)}" class="product ${a.products[p] ? 'yes' : 'no'}">${a.products[p] ? '✓' : '—'}</td>`).join('');
  const eventSummary = a => a.summitEvents.length ? a.summitEvents.map(e => `${e.date}: ${e.subject}`).join('; ') : 'No Salesforce Event found Sep 1–3';
  const rows = data.accounts.map(a => `
    <tr data-owner="${htmlEscape(a.referralOwner.toLowerCase())}" data-products="${htmlEscape(PRODUCTS.filter(p => a.products[p]).join(' ').toLowerCase())}" data-missing="${htmlEscape(a.missing.join(' ').toLowerCase())}" data-search="${htmlEscape([a.account, a.referralOwner, a.accountOwner, a.assignedAm, a.registeredCompanyNames.join(' '), a.registrants.map(r => r.name).join(' '), eventSummary(a)].join(' ').toLowerCase())}">
      <td data-label="Account" class="sticky"><strong>${htmlEscape(a.account)}</strong><span class="cohort">${htmlEscape(a.originalCohort)}</span><span>${htmlEscape(a.registeredCompanyNames.join(' / '))}</span></td>
      <td data-label="Owners"><strong>Referral:</strong> ${htmlEscape(a.referralOwner || 'Unassigned')}<span>Account: ${htmlEscape(a.accountOwner || 'Unassigned')}</span><span>AM: ${htmlEscape(a.assignedAm || 'Unassigned')}</span></td>
      <td data-label="Events Sep 1–3">${a.summitEvents.length ? `<strong>${a.summitEvents.length} scheduled</strong>` : '<em>None found</em>'}<span>${htmlEscape(eventSummary(a))}</span></td>
      <td data-label="Attendees" class="num">${a.attendeeCount}<span>${htmlEscape(a.registrants.map(r => r.name).join('; '))}</span></td>
      ${productCells(a)}
      <td data-label="Missing / target" class="missing">${a.missing.length ? a.missing.map(p => `<b>${htmlEscape(p)}</b>`).join(' ') : '<em>Complete set</em>'}</td>
      <td data-label="Details"><button class="details" type="button">View</button></td>
    </tr>
    <tr class="detail-row"><td colspan="${5 + PRODUCTS.length}"><div class="details-box">
      <div><strong>Registrant targeting notes</strong><ul>${a.registrants.map(r => `<li>${htmlEscape(r.name)}${r.title ? ` — ${htmlEscape(r.title)}` : ''}${r.department ? ` · ${htmlEscape(r.department)}` : ''}${r.referral ? ` · referred by ${htmlEscape(r.referral)}` : ''}${r.attendedBefore ? ` · attended before: ${htmlEscape(r.attendedBefore)}` : ''}</li>`).join('')}</ul><strong>Salesforce Events Sep 1–3</strong><ul>${a.summitEvents.length ? a.summitEvents.map(e => `<li>${htmlEscape(e.date)} — ${htmlEscape(e.subject)}${e.owner ? ` · owner: ${htmlEscape(e.owner)}` : ''}</li>`).join('') : '<li>No Salesforce Event found.</li>'}</ul></div>
      <div><strong>Owners</strong><p>Referral: ${htmlEscape(a.referralOwner || 'Unassigned')}<br>Account: ${htmlEscape(a.accountOwner || 'Unassigned')}<br>AM: ${htmlEscape(a.assignedAm || 'Unassigned')}</p><strong>Source product field</strong><p>${htmlEscape(a.productsRaw || '—')}</p><strong>Salesforce account</strong><p>${htmlEscape(a.sfAccountName || '—')}</p><strong>Match score</strong><p>${htmlEscape(a.matchScore)}</p></div>
    </div></td></tr>`).join('');

  const rollupRows = data.rollup.map(r => `<tr><td>${htmlEscape(r.owner)}</td><td class="num">${r.accounts}</td><td class="num">${r.attendees}</td><td class="num">${r.missingOpportunities}</td></tr>`).join('');

  const unmatchedRows = data.unmatched.map(u => `<tr><td>${htmlEscape(u.company)}</td><td>${htmlEscape(u.name)}</td><td>${htmlEscape(u.referral || '—')}</td></tr>`).join('');

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
.panel{background:var(--white);border:1px solid var(--border);border-radius:20px;margin-top:18px;padding:18px;box-shadow:0 10px 30px rgba(29,55,86,.06)}.panel h2{color:var(--navy);margin:0 0 12px;font-size:21px}.controls{display:grid;grid-template-columns:2fr repeat(3,1fr);gap:10px;margin-bottom:14px}.controls input,.controls select{border:1px solid var(--border);border-radius:12px;padding:11px 12px;font:inherit;color:var(--body);background:white}.actions{display:flex;justify-content:space-between;gap:10px;align-items:center;flex-wrap:wrap}.download{display:inline-block;background:var(--green);color:#113024;text-decoration:none;font-weight:800;border-radius:12px;padding:10px 14px}.note{font-size:12px;color:#647386}.table-wrap{overflow-x:auto;overflow-y:visible;-webkit-overflow-scrolling:touch;border:1px solid var(--border);border-radius:16px}table{width:100%;border-collapse:separate;border-spacing:0;min-width:980px}th{position:sticky;top:0;background:#eef6f8;color:var(--navy);font-size:12px;text-align:left;text-transform:uppercase;letter-spacing:.04em;padding:12px;border-bottom:1px solid var(--border);z-index:2}td{padding:12px;border-bottom:1px solid var(--border);vertical-align:top;font-size:13px}td span{display:block;color:#6b7787;font-size:11px;margin-top:3px;max-width:320px}.sticky{position:sticky;left:0;background:white;z-index:1;box-shadow:1px 0 0 var(--border)}.num{text-align:right;font-family:'Montserrat';font-weight:700;color:var(--navy)}.product{text-align:center;font:800 18px/1 'Montserrat'}.product.yes{color:var(--green)}.product.no{color:#b6c0cb}.cohort{display:inline-block!important;width:max-content;margin:7px 0 2px!important;background:var(--navy);color:var(--white)!important;border-radius:999px;padding:4px 9px;font-size:10px!important;font-weight:800;text-transform:uppercase;letter-spacing:.04em}.missing b{display:inline-block;margin:0 4px 4px 0;background:#eaf6ea;color:#2d6530;border:1px solid #cce8c8;border-radius:999px;padding:4px 8px;font-size:11px}.missing em{color:#7a8795}.details{border:0;background:var(--teal);color:white;border-radius:10px;padding:7px 11px;font-weight:800;cursor:pointer}.detail-row{display:none}.detail-row.open{display:table-row}.details-box{display:grid;grid-template-columns:2fr 1fr;gap:20px;background:#f8fbfc;border-radius:14px;padding:14px}.details-box ul{margin:8px 0 0;padding-left:18px}.details-box p{margin:6px 0 12px}.rollup{max-width:680px;min-width:520px}.warning{border-left:5px solid var(--teal);background:#f6fbfc}.footer{margin:22px 0;color:#6b7787;font-size:12px}
@media(max-width:900px){main{padding:0 12px}.hero{padding:28px 18px 42px}.stats{grid-template-columns:repeat(2,1fr)}.controls{grid-template-columns:1fr}.details-box{grid-template-columns:1fr}.table-wrap{overflow:visible;border:0}#accounts{min-width:0;border-spacing:0 12px}#accounts thead{display:none}#accounts tbody,#accounts tr,#accounts td{display:block;width:100%}#accounts tr:not(.detail-row){background:white;border:1px solid var(--border);border-radius:16px;padding:10px;box-shadow:0 8px 22px rgba(29,55,86,.07)}#accounts td{border:0;padding:8px 10px}#accounts td::before{content:attr(data-label);display:block;margin-bottom:3px;color:#647386;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.06em}.sticky{position:static;box-shadow:none}.num{text-align:left}.product{text-align:left;display:inline-block!important;width:auto!important;min-width:31%;font-size:16px}.missing b{margin-top:2px}.detail-row.open{display:block;background:white;border:1px solid var(--border);border-radius:16px;margin-top:-8px}.detail-row td{padding:10px}.rollup{min-width:0}.rollup table{min-width:520px}}
</style>
</head>
<body>
<div class="topbar"></div>
<header class="hero"><div class="wrap"><div class="eyebrow">Rev.io Summit 2026 · Client Targeting</div><h1>Client target dashboard</h1><p>Account-level view of Summit attendees who used non-sponsor/non-REVII discount codes, matched to the Master Client List products so referrers know which clients they own and which products to target onsite.</p><div class="meta"><span class="pill">Excludes REVII</span><span class="pill">Excludes SUMMITSPONSOR</span><span class="pill">Targets: Billing · New Rev.io · Payments</span><span class="pill">Cohort from Salesforce</span><span class="pill">Generated ${htmlEscape(new Date(data.generatedAt).toLocaleString('en-US', { timeZone: 'UTC', dateStyle: 'medium', timeStyle: 'short' }))} UTC</span></div></div></header>
<main>
<section class="stats">
${stat('Target accounts', data.summary.targetAccounts)}${stat('Matched attendees', data.summary.matchedClientRegistrants, `${data.summary.discountedNonSponsorRegistrants} discounted non-sponsor registrants`)}${stat('Referrers / owners', data.summary.referrers)}${stat('Accounts missing products', data.summary.accountsWithMissingProducts)}${stat('Missing product opps', data.summary.missingProductOpportunities)}${stat('Unmatched registrants', data.summary.unmatchedRegistrants, `${data.summary.ignoredSponsorRegistrants} sponsor/speaker records ignored`)}
</section>
<section class="panel"><div class="actions"><div><h2>Target account list</h2><div class="note">Owners show Referral Owner, Salesforce Account Owner, and AM Owner. Product checks come from Master Client List “Rev.io Product”; New Rev.io maps to PSA Web. Original cohort is pulled from Salesforce; Tigerpaw cohort is driven by PSA Account Status. Meetings are Salesforce Events dated Sep 1–3, 2026.</div></div><a class="download" href="assets/data/summit-target-accounts.csv">Download CSV</a></div>
<div class="controls"><input id="search" placeholder="Search account, attendee, owner…"><select id="owner"><option value="">All owners</option>${data.rollup.map(r => `<option>${htmlEscape(r.owner)}</option>`).join('')}</select><select id="missing"><option value="">All missing products</option>${PRODUCTS.map(p => `<option>${htmlEscape(p)}</option>`).join('')}<option value="none">No missing products</option></select><select id="have"><option value="">All current products</option>${PRODUCTS.map(p => `<option>${htmlEscape(p)}</option>`).join('')}</select></div>
<div class="table-wrap"><table id="accounts"><thead><tr><th class="sticky">Account</th><th>Owners</th><th>Events Sep 1–3</th><th>Attendees</th>${productHeader}<th>Missing / target</th><th>Details</th></tr></thead><tbody>${rows}</tbody></table></div></section>
<section class="panel"><h2>Referrer workload</h2><div class="table-wrap rollup"><table><thead><tr><th>Owner</th><th>Accounts</th><th>Attendees</th><th>Missing product opps</th></tr></thead><tbody>${rollupRows}</tbody></table></div></section>
<section class="panel warning"><h2>Unmatched discounted registrants</h2><p class="note">These passed the discount-code rule but did not confidently match a Master Client List account. They are excluded from target account stats until manually mapped.</p><div class="table-wrap"><table><thead><tr><th>Company</th><th>Registrant</th><th>Referral</th></tr></thead><tbody>${unmatchedRows || '<tr><td colspan="3">No unmatched registrants.</td></tr>'}</tbody></table></div></section>
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
