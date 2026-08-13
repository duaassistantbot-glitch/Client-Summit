const fs = require('fs');
const path = require('path');
const { createRequire } = require('module');
const ExcelJS = require('exceljs');

const outDir = path.resolve(__dirname, '..');
const workspaceDir = path.resolve(outDir, '..');
const workspaceRequire = createRequire(path.join(workspaceDir, 'gmail-tool', 'package.json'));
const XLSX = workspaceRequire('xlsx');

const currentPath = path.resolve(process.env.CURRENT_XLSX || path.join(outDir, 'assets', 'source', 'registration-report-deduped-2026-08-12.xlsx'));
const incomingPath = process.env.INCOMING_XLSX ? path.resolve(process.env.INCOMING_XLSX) : '';
const outputPath = path.resolve(process.env.OUT_XLSX || currentPath);

function normEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function plainCellValue(value) {
  if (value && typeof value === 'object' && 'text' in value) return value.text;
  if (value && typeof value === 'object' && 'result' in value) return value.result;
  if (value instanceof Date) return value;
  return value ?? '';
}

function readIncomingRows(file) {
  const workbook = XLSX.readFile(file, { raw: false, cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

(async () => {
  if (!incomingPath) {
    throw new Error('Set INCOMING_XLSX=/path/to/latest-registration-export.xlsx-or-csv');
  }
  if (!fs.existsSync(currentPath)) throw new Error(`Current workbook not found: ${currentPath}`);
  if (!fs.existsSync(incomingPath)) throw new Error(`Incoming workbook not found: ${incomingPath}`);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(currentPath);
  const worksheet = workbook.worksheets[0];
  const headers = worksheet.getRow(1).values.slice(1).map((value) => String(value || '').trim());
  const emailCol = headers.indexOf('Email') + 1;
  if (!emailCol) throw new Error('Email header not found in current workbook');

  const existingEmails = new Set();
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const email = normEmail(plainCellValue(row.getCell(emailCol).value));
    if (email) existingEmails.add(email);
  });

  const incomingRows = readIncomingRows(incomingPath);
  const added = [];
  const skippedExisting = [];
  const skippedCancelled = [];
  const skippedNoEmail = [];

  for (const row of incomingRows) {
    const email = normEmail(row.Email);
    const status = String(row.Status || '');
    const name = `${row['First Name (1557966)'] || ''} ${row['Last Name (1557967)'] || ''}`.trim();

    if (!email) {
      skippedNoEmail.push(name || '(blank name)');
      continue;
    }
    if (/cancel/i.test(status)) {
      skippedCancelled.push({ name, email, status });
      continue;
    }
    if (existingEmails.has(email)) {
      skippedExisting.push({ name, email });
      continue;
    }

    worksheet.addRow(headers.map((header) => row[header] ?? ''));
    existingEmails.add(email);
    added.push({ name, email, company: row.Company || '', dateRegistered: row['Date Registered'] || '' });
  }

  if (outputPath === currentPath) {
    const backupPath = currentPath.replace(/\.xlsx$/i, `.pre-append-${Date.now()}.xlsx`);
    fs.copyFileSync(currentPath, backupPath);
    console.error(`Backup written: ${backupPath}`);
  }

  await workbook.xlsx.writeFile(outputPath);
  console.log(JSON.stringify({
    currentPath,
    incomingPath,
    outputPath,
    existingBefore: existingEmails.size - added.length,
    incomingRows: incomingRows.length,
    added: added.length,
    skippedExisting: skippedExisting.length,
    skippedCancelled: skippedCancelled.length,
    skippedNoEmail: skippedNoEmail.length,
    addedPeople: added
  }, null, 2));
})();
