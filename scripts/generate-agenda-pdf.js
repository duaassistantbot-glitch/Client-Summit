const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');
const { AGENDA_DAYS } = require('./agenda-data');

const outDir = path.resolve(__dirname, '..');
const assetsDir = path.join(outDir, 'assets');
const htmlPath = path.join(assetsDir, 'Revio-Summit-2026-Agenda.html');
const pdfPath = path.join(assetsDir, 'Revio-Summit-2026-Agenda.pdf');

function safe(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  })[char]);
}

function sessionClass(type) {
  return String(type || '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'session';
}

function findChrome() {
  const candidates = ['google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser'];
  for (const command of candidates) {
    try {
      const resolved = execFileSync('which', [command], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim();
      if (resolved) return resolved;
    } catch {
      // Try the next Chrome-compatible binary.
    }
  }
  throw new Error('No Chrome-compatible browser found for PDF generation.');
}

function renderAgendaHtml() {
  const allSessions = AGENDA_DAYS.flatMap((day) => day.sessions);

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Rev.io Summit 2026 Agenda</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
@page { size: Letter; margin: .42in .38in .44in; }
* { box-sizing: border-box; }
html, body { margin: 0; }
body {
  background: #f5f7fa;
  color: #1D3756;
  font-family: "Open Sans", Arial, sans-serif;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
  font-size: 10px;
}
.cover {
  position: relative;
  overflow: hidden;
  min-height: 1.65in;
  margin: -.42in -.38in .22in;
  padding: .34in .38in .22in;
  background:
    linear-gradient(115deg, rgba(52,189,229,.23), transparent 36%),
    linear-gradient(250deg, rgba(198,241,120,.16), transparent 42%),
    linear-gradient(180deg, #0a141f 0%, #0c1925 52%, #060e18 100%);
}
.hero-art {
  position: absolute;
  inset: 0;
  z-index: 0;
  opacity: .62;
}
.content { position: relative; z-index: 1; }
.topbar {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: .18in;
  align-items: end;
}
h1 {
  margin: 0;
  color: #FFFFFF;
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 30px;
  line-height: 1;
  letter-spacing: 0;
  text-transform: uppercase;
}
.subtitle {
  margin-top: .05in;
  color: #FFFFFF;
  font-size: 10.5px;
  font-weight: 800;
  letter-spacing: 1.2px;
  text-transform: uppercase;
}
.meta {
  display: grid;
  gap: .045in;
  justify-items: end;
}
.meta span, .type {
  border: 1px solid rgba(127,217,239,.36);
  border-radius: 6px;
  background: rgba(52,189,229,.12);
  color: #FFFFFF;
  font-weight: 800;
  text-transform: uppercase;
}
.meta span {
  padding: .045in .075in;
  font-size: 7.5px;
  letter-spacing: .7px;
}
.note {
  margin-top: .13in;
  max-width: 6.6in;
  color: #FFFFFF;
  font-size: 9px;
  font-weight: 700;
  line-height: 1.35;
}
.day {
  margin: 0 0 .18in;
  border: 1px solid #DDE2E8;
  border-radius: 7px;
  overflow: hidden;
  background: #FFFFFF;
  break-inside: avoid;
  page-break-inside: avoid;
}
.day.allow-break {
  break-inside: auto;
  page-break-inside: auto;
}
.day-header {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: .14in;
  align-items: center;
  padding: .105in .13in .1in;
  background: linear-gradient(90deg, #1D3756, #2399B5);
}
.day-label {
  color: #FFFFFF;
  font-size: 7px;
  font-weight: 800;
  letter-spacing: 1.1px;
  text-transform: uppercase;
}
.day-date {
  margin-top: .015in;
  color: #FFFFFF;
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 14px;
  font-weight: 800;
}
.day-theme {
  color: #FFFFFF;
  font-size: 8px;
  font-weight: 800;
  letter-spacing: .5px;
  text-transform: uppercase;
}
.sessions { padding: 0; }
.session {
  display: grid;
  grid-template-columns: 1.05in minmax(0, 1fr);
  gap: .12in;
  align-items: start;
  padding: .074in .13in .07in;
  border-bottom: 1px solid #DDE2E8;
  break-inside: avoid;
  page-break-inside: avoid;
}
.session:last-child { border-bottom: 0; }
.time {
  color: #1D3756;
  font-size: 8.5px;
  font-weight: 800;
  line-height: 1.25;
  font-variant-numeric: tabular-nums;
  text-transform: uppercase;
}
.title {
  color: #1D3756;
  font-size: 9.2px;
  font-weight: 800;
  line-height: 1.24;
}
.type {
  display: inline-block;
  justify-self: start;
  padding: .036in .052in;
  border-color: rgba(35,153,181,.28);
  background: rgba(35,153,181,.11);
  color: #1D3756;
  font-size: 6.6px;
  line-height: 1.1;
  letter-spacing: .35px;
}
.type.keynote, .type.general-session {
  border-color: rgba(110,190,79,.45);
  background: rgba(110,190,79,.14);
}
.type.networking, .type.registration {
  border-color: rgba(29,55,86,.24);
  background: rgba(29,55,86,.08);
}
.footer {
  margin-top: .1in;
  display: flex;
  justify-content: space-between;
  color: #3d4d5c;
  font-size: 7.5px;
  font-weight: 700;
  border-top: 1px solid rgba(127,217,239,.18);
  padding-top: .07in;
}
@media print {
  .day:nth-of-type(n+3) {
    break-inside: auto;
    page-break-inside: auto;
  }
}
</style>
</head>
<body>
<section class="cover">
  <svg class="hero-art" viewBox="0 0 1100 850" preserveAspectRatio="none" aria-hidden="true">
    <defs>
      <linearGradient id="summitSky" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#060e18"/><stop offset=".52" stop-color="#15283f"/><stop offset="1" stop-color="#2399B5"/></linearGradient>
      <linearGradient id="aurora" x1="0" y1="0" x2="1" y2="0"><stop stop-color="#34bde5" stop-opacity=".03"/><stop offset=".48" stop-color="#4fd1c5" stop-opacity=".34"/><stop offset="1" stop-color="#c6f178" stop-opacity=".13"/></linearGradient>
      <linearGradient id="farRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#294b72"/><stop offset="1" stop-color="#0f1b2a"/></linearGradient>
      <linearGradient id="nearRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#1d3756"/><stop offset="1" stop-color="#060e18"/></linearGradient>
    </defs>
    <rect width="1100" height="850" fill="url(#summitSky)"/>
    <path d="M-30 102 C162 14 314 120 484 58 C662 -8 806 70 1130 16 L1130 276 C932 314 756 244 590 302 C400 368 220 260 -30 394Z" fill="url(#aurora)" opacity=".9"/>
    <path d="M0 662 L124 578 L230 618 L362 504 L492 596 L620 468 L772 624 L912 520 L1100 594 L1100 850 L0 850Z" fill="url(#farRidge)" opacity=".72"/>
    <path d="M0 738 L152 660 L286 704 L426 604 L558 698 L704 580 L846 714 L982 634 L1100 690 L1100 850 L0 850Z" fill="url(#nearRidge)" opacity=".82"/>
  </svg>
  <div class="content topbar">
    <div>
      <h1>Rev.io Summit 2026 Agenda</h1>
      <div class="subtitle">Prepare for Tomorrow</div>
      <div class="note">Updated team-share agenda with session titles and date times only.</div>
    </div>
    <div class="meta">
      <span>Sep 1-3, 2026</span>
      <span>Atlanta, GA</span>
      <span>${allSessions.length} Sessions</span>
    </div>
  </div>
</section>
<main>
    ${AGENDA_DAYS.map((day, index) => `
      <article class="day ${index > 0 ? 'allow-break' : ''}">
        <div class="day-header">
          <div>
            <div class="day-label">${safe(day.label)}</div>
            <div class="day-date">${safe(day.date)}</div>
          </div>
          <div class="day-theme">${safe(day.theme)}</div>
        </div>
        <div class="sessions">
          ${day.sessions.map((session) => `
            <div class="session">
              <div class="time">${safe(session.start)}<br>to ${safe(session.end)}</div>
              <div>
                <div class="title">${safe(session.title)}</div>
              </div>
            </div>
          `).join('')}
        </div>
      </article>
    `).join('')}
  <div class="footer"><span>Rev.io Summit 2026</span><span>Titles and date times only</span></div>
</main>
</body>
</html>`;
}

fs.mkdirSync(assetsDir, { recursive: true });
fs.writeFileSync(htmlPath, renderAgendaHtml());

const chrome = findChrome();
execFileSync(chrome, [
  '--headless=new',
  '--disable-gpu',
  '--no-sandbox',
  '--print-to-pdf=' + pdfPath,
  '--print-to-pdf-no-header',
  'file://' + htmlPath
], { stdio: 'inherit' });

console.log(`Generated ${pdfPath}`);
