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

function renderAgendaHtml() {
  const allSessions = AGENDA_DAYS.flatMap((day) => day.sessions);
  const breakoutCount = allSessions.filter((session) => session.type === 'Breakout' || session.type === 'Workshop').length;
  const networkingCount = allSessions.filter((session) => /Networking|Meal|Expo|Break/i.test(`${session.type} ${session.title}`)).length;

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Rev.io Summit 2026 Agenda</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600;700;800&display=swap" rel="stylesheet">
<style>
@page { size: Letter; margin: 0; }
* { box-sizing: border-box; }
body {
  margin: 0;
  background: #0a141f;
  color: #FFFFFF;
  font-family: "Open Sans", Arial, sans-serif;
  -webkit-print-color-adjust: exact;
  print-color-adjust: exact;
}
.page {
  position: relative;
  width: 8.5in;
  min-height: 11in;
  overflow: hidden;
  background:
    linear-gradient(115deg, rgba(52,189,229,.16), transparent 34%),
    linear-gradient(250deg, rgba(198,241,120,.12), transparent 38%),
    linear-gradient(180deg, #0a141f 0%, #0c1925 50%, #060e18 100%);
  page-break-after: always;
}
.page:last-child { page-break-after: auto; }
.cover {
  display: flex;
  flex-direction: column;
  justify-content: space-between;
  padding: .58in .62in .5in;
}
.hero-art {
  position: absolute;
  inset: 0;
  opacity: .98;
  z-index: 0;
}
.content { position: relative; z-index: 1; }
.eyebrow {
  display: inline-flex;
  align-items: center;
  gap: .1in;
  border: 1px solid rgba(127,217,239,.42);
  color: #FFFFFF;
  background: rgba(52,189,229,.12);
  padding: .07in .13in;
  border-radius: 999px;
  font-size: 9px;
  letter-spacing: 1.8px;
  text-transform: uppercase;
  font-weight: 800;
}
.eyebrow::before {
  content: "";
  width: .07in;
  height: .07in;
  border-radius: 999px;
  background: #34bde5;
  box-shadow: 0 0 0 .04in rgba(52,189,229,.18);
}
h1, h2, h3 { font-family: Montserrat, "Open Sans", sans-serif; margin: 0; }
h1 {
  margin-top: .22in;
  max-width: 6.8in;
  font-size: 62px;
  line-height: .94;
  letter-spacing: 0;
  text-transform: uppercase;
}
.subtitle {
  margin-top: .16in;
  font-size: 22px;
  font-weight: 800;
}
.meta {
  display: flex;
  flex-wrap: wrap;
  gap: .1in;
  margin-top: .24in;
}
.meta span {
  border: 1px solid rgba(255,255,255,.16);
  border-radius: 999px;
  background: rgba(10,20,31,.46);
  padding: .08in .12in;
  color: #FFFFFF;
  font-size: 10px;
  font-weight: 800;
  letter-spacing: 1.2px;
  text-transform: uppercase;
}
.cover-grid {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: .14in;
  margin-top: .36in;
}
.stat {
  min-height: .95in;
  padding: .16in;
  border: 1px solid rgba(255,255,255,.12);
  border-radius: 8px;
  background: linear-gradient(145deg, rgba(21,40,63,.94), rgba(10,20,31,.78));
}
.stat-label {
  color: #7fd9ef;
  font-size: 8px;
  font-weight: 800;
  letter-spacing: 1.2px;
  text-transform: uppercase;
}
.stat-value {
  margin-top: .06in;
  color: #FFFFFF;
  font-family: Montserrat, "Open Sans", sans-serif;
  font-size: 29px;
  font-weight: 800;
}
.cover-note {
  max-width: 6.9in;
  color: #FFFFFF;
  font-size: 13px;
  line-height: 1.55;
}
.stripe {
  width: 1.65in;
  height: .04in;
  margin: .18in 0 .2in;
  background: linear-gradient(90deg, #2399B5 0%, #6EBE4F 100%);
}
.agenda-page { padding: .44in .46in .4in; }
.page-head {
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: .25in;
  padding-bottom: .17in;
  border-bottom: 1px solid rgba(127,217,239,.24);
}
.day-label {
  color: #7fd9ef;
  font-size: 9px;
  font-weight: 800;
  letter-spacing: 1.6px;
  text-transform: uppercase;
}
.day-title {
  margin-top: .04in;
  color: #FFFFFF;
  font-size: 24px;
  font-weight: 800;
}
.theme {
  color: #FFFFFF;
  font-size: 11px;
  font-weight: 800;
  text-align: right;
}
.session-list { margin-top: .16in; }
.session {
  display: grid;
  grid-template-columns: .82in minmax(0, 1fr) 1.28in;
  gap: .13in;
  padding: .1in 0;
  border-bottom: 1px solid rgba(255,255,255,.08);
  break-inside: avoid;
}
.time {
  color: #FFFFFF;
  font-size: 9px;
  font-weight: 800;
  line-height: 1.3;
  font-variant-numeric: tabular-nums;
}
.title {
  color: #FFFFFF;
  font-size: 10.5px;
  font-weight: 800;
  line-height: 1.33;
}
.details {
  margin-top: .03in;
  color: #b9c7d6;
  font-size: 8.5px;
  font-weight: 700;
  line-height: 1.35;
}
.room {
  display: flex;
  align-items: start;
  justify-content: flex-end;
  gap: .06in;
  text-align: right;
}
.type {
  display: inline-block;
  max-width: 1.08in;
  border: 1px solid rgba(127,217,239,.26);
  border-radius: 999px;
  background: rgba(52,189,229,.12);
  padding: .035in .075in;
  color: #FFFFFF;
  font-size: 7px;
  font-weight: 800;
  line-height: 1.15;
  text-transform: uppercase;
}
.type.break, .type.meal-expo { border-color: rgba(198,241,120,.35); background: rgba(110,190,79,.12); }
.footer {
  position: absolute;
  left: .46in;
  right: .46in;
  bottom: .22in;
  display: flex;
  justify-content: space-between;
  color: #b9c7d6;
  font-size: 8px;
  border-top: 1px solid rgba(127,217,239,.18);
  padding-top: .08in;
}
</style>
</head>
<body>
<section class="page cover">
  <svg class="hero-art" viewBox="0 0 850 1100" preserveAspectRatio="none" aria-hidden="true">
    <defs>
      <linearGradient id="summitSky" x1="0" y1="0" x2="1" y2="1"><stop stop-color="#060e18"/><stop offset=".52" stop-color="#15283f"/><stop offset="1" stop-color="#2399B5"/></linearGradient>
      <linearGradient id="aurora" x1="0" y1="0" x2="1" y2="0"><stop stop-color="#34bde5" stop-opacity=".03"/><stop offset=".48" stop-color="#4fd1c5" stop-opacity=".38"/><stop offset="1" stop-color="#c6f178" stop-opacity=".14"/></linearGradient>
      <linearGradient id="farRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#294b72"/><stop offset="1" stop-color="#0f1b2a"/></linearGradient>
      <linearGradient id="nearRidge" x1="0" y1="0" x2="0" y2="1"><stop stop-color="#1d3756"/><stop offset="1" stop-color="#060e18"/></linearGradient>
    </defs>
    <rect width="850" height="1100" fill="url(#summitSky)"/>
    <path d="M-30 142 C120 24 244 152 386 78 C526 6 650 92 880 22 L880 366 C716 410 600 320 466 386 C310 462 182 330 -30 486Z" fill="url(#aurora)" opacity=".92"/>
    <path d="M0 760 L78 690 L146 718 L235 612 L318 690 L412 570 L520 724 L628 545 L730 706 L850 650 L850 1100 L0 1100Z" fill="url(#farRidge)" opacity=".88"/>
    <path d="M0 874 L102 794 L192 832 L306 730 L420 812 L536 690 L668 834 L760 748 L850 798 L850 1100 L0 1100Z" fill="url(#nearRidge)" opacity=".97"/>
    <path d="M0 874 L102 794 L192 832 L306 730 L420 812 L536 690 L668 834 L760 748 L850 798" fill="none" stroke="#7fd9ef" stroke-width="1.4" opacity=".35"/>
  </svg>
  <div class="content">
    <div class="eyebrow">Team Agenda</div>
    <h1>Rev.io Summit 2026</h1>
    <div class="subtitle">Prepare for Tomorrow.</div>
    <div class="stripe"></div>
    <p class="cover-note">A shareable team agenda for three days of market perspective, Rev.io platform strategy, AI-enabled operations, breakout learning, sponsor sessions, and customer networking.</p>
    <div class="meta"><span>September 1-3, 2026</span><span>Atlanta, GA</span><span>Client Summit Program</span></div>
  </div>
  <div class="content cover-grid">
    <div class="stat"><div class="stat-label">Program Days</div><div class="stat-value">${AGENDA_DAYS.length}</div></div>
    <div class="stat"><div class="stat-label">Agenda Blocks</div><div class="stat-value">${allSessions.length}</div></div>
    <div class="stat"><div class="stat-label">Breakouts + Workshops</div><div class="stat-value">${breakoutCount}</div></div>
    <div class="stat"><div class="stat-label">Networking Moments</div><div class="stat-value">${networkingCount}</div></div>
    <div class="stat"><div class="stat-label">Main Room</div><div class="stat-value">GB 3&4</div></div>
    <div class="stat"><div class="stat-label">Offsite Event</div><div class="stat-value" style="font-size:21px;">Truist Park</div></div>
  </div>
</section>
${AGENDA_DAYS.map((day, dayIndex) => `
<section class="page agenda-page">
  <div class="page-head">
    <div>
      <div class="day-label">${safe(day.label)}</div>
      <h2 class="day-title">${safe(day.date)}</h2>
    </div>
    <div class="theme">${safe(day.theme)}</div>
  </div>
  <div class="session-list">
    ${day.sessions.map((session) => `
      <div class="session">
        <div class="time">${safe(session.start)}<br>${safe(session.end)}</div>
        <div>
          <div class="title">${safe(session.title)}</div>
          <div class="details">${safe(session.track)}</div>
        </div>
        <div class="room">
          <span class="type ${sessionClass(session.type)}">${safe(session.type)}</span>
          <div class="details">${safe(session.room)}</div>
        </div>
      </div>
    `).join('')}
  </div>
  <div class="footer"><span>Rev.io Summit 2026</span><span>${dayIndex + 2}</span></div>
</section>
`).join('')}
</body>
</html>`;
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
