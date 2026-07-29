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
@page { size: Letter landscape; margin: 0; }
* { box-sizing: border-box; }
html, body { margin: 0; width: 11in; min-height: 8.5in; }
body {
  background: #0a141f;
  color: #FFFFFF;
  font-family: "Open Sans", Arial, sans-serif;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.page {
  position: relative;
  width: 11in;
  height: 8.5in;
  overflow: hidden;
  padding: .28in .32in .24in;
  background:
    linear-gradient(115deg, rgba(52,189,229,.16), transparent 34%),
    linear-gradient(250deg, rgba(198,241,120,.12), transparent 38%),
    linear-gradient(180deg, #0a141f 0%, #0c1925 52%, #060e18 100%);
}
.hero-art {
  position: absolute;
  inset: 0;
  z-index: 0;
  opacity: .74;
}
.content { position: relative; z-index: 1; }
.topbar {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  gap: .22in;
  align-items: end;
  padding-bottom: .12in;
  border-bottom: 1px solid rgba(127,217,239,.24);
}
h1 {
  margin: 0;
  color: #FFFFFF;
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 30px;
  line-height: .95;
  letter-spacing: 0;
  text-transform: uppercase;
}
.subtitle {
  margin-top: .04in;
  color: #FFFFFF;
  font-size: 10px;
  font-weight: 800;
  letter-spacing: 1.2px;
  text-transform: uppercase;
}
.meta {
  display: flex;
  gap: .06in;
  justify-content: flex-end;
  flex-wrap: wrap;
}
.meta span, .type {
  border: 1px solid rgba(127,217,239,.28);
  border-radius: 999px;
  background: rgba(52,189,229,.12);
  color: #FFFFFF;
  font-weight: 800;
  text-transform: uppercase;
}
.meta span {
  padding: .045in .075in;
  font-size: 6.8px;
  letter-spacing: .7px;
}
.agenda-grid {
  display: grid;
  grid-template-columns: .65fr 1.35fr 1fr;
  gap: .12in;
  margin-top: .13in;
}
.day {
  min-height: 7.08in;
  border: 1px solid rgba(255,255,255,.11);
  border-radius: 8px;
  overflow: hidden;
  background: rgba(10,20,31,.7);
}
.day-header {
  min-height: .48in;
  padding: .09in .11in .08in;
  background: linear-gradient(90deg, rgba(29,55,86,.98), rgba(35,153,181,.28));
  border-bottom: 1px solid rgba(127,217,239,.22);
}
.day-label {
  color: #7fd9ef;
  font-size: 6.6px;
  font-weight: 800;
  letter-spacing: 1px;
  text-transform: uppercase;
}
.day-date {
  margin-top: .02in;
  color: #FFFFFF;
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 11.2px;
  font-weight: 800;
}
.day-theme {
  margin-top: .018in;
  color: #FFFFFF;
  font-size: 6.9px;
  font-weight: 800;
}
.sessions { padding: .055in .085in .07in; }
.session {
  display: grid;
  grid-template-columns: .52in minmax(0, 1fr);
  gap: .055in;
  padding: .031in 0;
  border-bottom: 1px solid rgba(255,255,255,.075);
}
.session:last-child { border-bottom: 0; }
.time {
  color: #FFFFFF;
  font-size: 6.25px;
  font-weight: 800;
  line-height: 1.15;
  font-variant-numeric: tabular-nums;
}
.title {
  color: #FFFFFF;
  font-size: 6.35px;
  font-weight: 800;
  line-height: 1.17;
}
.detail {
  display: flex;
  gap: .035in;
  align-items: center;
  margin-top: .018in;
  color: #b9c7d6;
  font-size: 5.25px;
  font-weight: 700;
  line-height: 1.1;
}
.type {
  flex: none;
  padding: .018in .038in;
  font-size: 4.6px;
  line-height: 1;
  letter-spacing: .25px;
}
.type.break, .type.meal-expo {
  border-color: rgba(198,241,120,.35);
  background: rgba(110,190,79,.12);
}
.type.networking { border-color: rgba(198,241,120,.38); }
.room {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.footer {
  position: absolute;
  left: .32in;
  right: .32in;
  bottom: .09in;
  display: flex;
  justify-content: space-between;
  color: #b9c7d6;
  font-size: 6.5px;
  border-top: 1px solid rgba(127,217,239,.18);
  padding-top: .045in;
}
</style>
</head>
<body>
<section class="page">
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
      <div class="subtitle">Prepare for Tomorrow.</div>
    </div>
    <div class="meta">
      <span>Sep 1-3, 2026</span>
      <span>Atlanta, GA</span>
      <span>${allSessions.length} Agenda Blocks</span>
    </div>
  </div>
  <div class="content agenda-grid">
    ${AGENDA_DAYS.map((day) => `
      <article class="day">
        <div class="day-header">
          <div class="day-label">${safe(day.label)}</div>
          <div class="day-date">${safe(day.date)}</div>
          <div class="day-theme">${safe(day.theme)}</div>
        </div>
        <div class="sessions">
          ${day.sessions.map((session) => `
            <div class="session">
              <div class="time">${safe(session.start)}<br>${safe(session.end)}</div>
              <div>
                <div class="title">${safe(session.title)}</div>
                <div class="detail">
                  <span class="type ${sessionClass(session.type)}">${safe(session.type)}</span>
                  <span class="room">${safe(session.track)} / ${safe(session.room)}</span>
                </div>
              </div>
            </div>
          `).join('')}
        </div>
      </article>
    `).join('')}
  </div>
  <div class="footer"><span>Rev.io Summit 2026</span><span>One-page team agenda</span></div>
</section>
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
