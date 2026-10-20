// Builds print materials: an A4 tri-fold brochure and an A4 one-page flyer (PDF).
// Run: node build-print.js   (needs react, react-dom, react-icons, playwright)
const path = require("path");
const fs = require("fs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const fa = require("react-icons/fa");
const { chromium } = require("playwright");

const OUT_DIR = path.join(__dirname, "print");
fs.mkdirSync(OUT_DIR, { recursive: true });

const ic = (Comp, color = "#ffffff") =>
  ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color, size: "100%" }));
const dot = (Comp, cls = "") => `<span class="dot ${cls}">${ic(Comp)}</span>`;

const CSS = `
:root { --navy:#1F2A6B; --red:#C8100A; --blue:#22428E; --sky:#5B7BC4; --mist:#E8ECF7; --tint:#F3F5FB; --ink:#1A1F36; --gray:#5F6577; }
* { box-sizing:border-box; margin:0; padding:0; }
html, body { background:#fff; }
body { font-family: Carlito, Calibri, Arial, sans-serif; color:var(--ink); -webkit-print-color-adjust:exact; print-color-adjust:exact; }
h1,h2,h3 { color:var(--navy); line-height:1.1; }
.eyebrow { font-size:8.5pt; font-weight:700; letter-spacing:.18em; text-transform:uppercase; color:var(--red); margin-bottom:2.5mm; }
.dot { display:inline-flex; flex:none; align-items:center; justify-content:center; width:10mm; height:10mm; border-radius:50%; background:var(--navy); padding:2.6mm; }
.dot.red { background:var(--red); } .dot.blue { background:var(--blue); } .dot.sm { width:8mm; height:8mm; padding:2mm; }
.dot svg { width:100%; height:100%; }
.row { display:flex; gap:3.5mm; align-items:flex-start; }
.row h3 { font-size:12pt; margin-bottom:.6mm; }
.row p { font-size:9.8pt; color:var(--gray); line-height:1.3; }
.logo-card { background:#fff; border-radius:3mm; padding:3.5mm 5mm; display:inline-block; }
.logo-card img { height:13mm; display:block; }
.water { position:absolute; opacity:.09; pointer-events:none; }
`;

const BROCHURE_CSS = `
@page { size: 297mm 210mm; margin:0; }
.page { width:297mm; height:210mm; display:flex; page-break-after:always; overflow:hidden; }
.page:last-child { page-break-after:auto; }
.panel { position:relative; height:210mm; padding:13mm 10mm 11mm; display:flex; flex-direction:column; }
.w97 { width:97mm; } .w100 { width:100mm; }
.dark { background:var(--navy); color:#fff; } .dark h1, .dark h2, .dark h3 { color:#fff; }
.tint { background:var(--tint); }
h2 { font-size:19pt; margin-bottom:5mm; }
.stack > * + * { margin-top:7.5mm; }
.stats { display:grid; grid-template-columns:1fr 1fr; gap:3.5mm; margin-top:8mm; }
.stat { background:var(--tint); border-radius:2.5mm; padding:4.5mm 4mm; }
.stat b { display:block; font-size:22pt; color:var(--red); line-height:1; }
.stat span { display:block; font-size:9.5pt; font-weight:700; color:var(--navy); margin-top:1.2mm; }
.stat small { display:block; font-size:8pt; color:var(--gray); margin-top:.6mm; line-height:1.25; }
.stat.hl { background:var(--navy); } .stat.hl b, .stat.hl span { color:#fff; } .stat.hl small { color:var(--mist); }
.lead { font-size:12pt; line-height:1.45; }
.market { background:#fff; border-radius:2.5mm; padding:4mm 4mm; }
.market + .market { margin-top:4mm; }
.market h3 { font-size:11pt; letter-spacing:.06em; display:flex; align-items:center; gap:2mm; }
.market h3 .pin { width:3.6mm; height:3.6mm; display:inline-flex; }
.market p { font-size:9.5pt; color:var(--ink); line-height:1.3; margin-top:.8mm; }
.market p em { font-style:normal; font-weight:700; color:var(--red); }
.promise { margin-top:auto; background:var(--navy); color:#fff; border-radius:2.5mm; padding:3.5mm 4mm; font-size:9.3pt; line-height:1.35; }
.promise b { color:#fff; display:block; font-size:10.5pt; margin-bottom:.8mm; }
.steps { counter-reset:s; }
.step { display:flex; gap:3.5mm; align-items:flex-start; }
.step + .step { margin-top:5mm; }
.serve { margin-top:9mm; } .serve div { display:flex; align-items:center; gap:3mm; font-size:11pt; font-weight:700; color:var(--navy); } .serve div + div { margin-top:3mm; }
.num { flex:none; width:8mm; height:8mm; border-radius:50%; background:var(--red); color:#fff; font-weight:700; font-size:11pt; display:flex; align-items:center; justify-content:center; }
.step h3 { font-size:11.5pt; } .step p { font-size:9.5pt; color:var(--gray); line-height:1.3; margin-top:.5mm; }
.contact { margin-top:auto; background:var(--navy); border-radius:3mm; padding:5mm; color:#fff; }
.contact .eyebrow { color:var(--sky); }
.contact .line { display:flex; align-items:center; gap:3mm; font-size:12.5pt; font-weight:700; }
.contact .line + .line { margin-top:3mm; }
.reg { font-size:7.5pt; color:var(--gray); margin-top:3mm; }
.cover h1 { font-size:30pt; margin-top:22mm; }
.cover .sub { font-size:12pt; color:var(--mist); margin-top:5mm; line-height:1.35; }
.cover .markets { margin-top:auto; }
.cover .markets div { display:flex; align-items:center; gap:2.5mm; font-size:11.5pt; font-weight:700; }
.cover .markets div + div { margin-top:2.6mm; }
.cover .markets .dot { width:7mm; height:7mm; padding:1.7mm; }
.cover .tag { margin-top:7mm; font-size:9.5pt; color:var(--sky); font-weight:700; letter-spacing:.12em; text-transform:uppercase; }
.svc { display:flex; gap:3.5mm; align-items:flex-start; }
.svc + .svc { margin-top:7mm; }
.svc h3 { font-size:12pt; } .svc p { font-size:9.8pt; color:var(--gray); line-height:1.3; margin-top:.5mm; }
.case { margin-top:auto; background:var(--tint); border-radius:2.5mm; padding:3.5mm 4mm; }
.case b { color:var(--red); font-size:10.5pt; display:block; margin-bottom:.8mm; }
.case p { font-size:9.2pt; line-height:1.35; }
`;

const FLYER_CSS = `
@page { size: 210mm 297mm; margin:0; }
.flyer { width:210mm; height:297mm; display:flex; flex-direction:column; overflow:hidden; }
.head { position:relative; background:var(--navy); color:#fff; padding:13mm 15mm 12mm; overflow:hidden; }
.head h1 { color:#fff; font-size:31pt; margin-top:9mm; max-width:150mm; }
.head p { font-size:13pt; color:var(--mist); margin-top:4mm; max-width:140mm; line-height:1.35; }
.body { padding:10mm 15mm 0; flex:1; display:flex; flex-direction:column; }
h2 { font-size:17pt; margin-bottom:4.5mm; }
.grid { display:grid; grid-template-columns:1fr 1fr; gap:5.5mm 8mm; }
.mk { display:grid; grid-template-columns:repeat(4,1fr); gap:3mm; margin-top:9mm; }
.mk div { background:var(--tint); border-radius:2.5mm; padding:3.5mm; }
.mk h3 { font-size:11pt; letter-spacing:.05em; }
.mk p { font-size:8.8pt; color:var(--gray); margin-top:1mm; line-height:1.3; }
.why { display:flex; gap:4mm; margin-top:9mm; }
.why div { flex:1; border:.3mm solid var(--mist); border-radius:2.5mm; padding:3.5mm; }
.why b { display:block; color:var(--navy); font-size:11pt; }
.why span { display:block; font-size:9pt; color:var(--gray); margin-top:.8mm; line-height:1.3; }
.foot { margin-top:auto; background:var(--navy); color:#fff; padding:7mm 15mm; display:flex; align-items:center; justify-content:space-between; }
.foot .cta { font-size:15pt; font-weight:700; }
.foot .cta small { display:block; font-size:9.5pt; color:var(--sky); font-weight:700; letter-spacing:.14em; text-transform:uppercase; margin-bottom:1mm; }
.foot .line { display:flex; align-items:center; gap:3mm; font-size:13pt; font-weight:700; }
.foot .line + .line { margin-top:2mm; }
.promise { margin-top:8mm; background:var(--tint); border-radius:2.5mm; padding:3.5mm 4.5mm; font-size:10pt; line-height:1.35; display:flex; gap:3.5mm; align-items:center; }
.promise b { color:var(--navy); }
`;

const MARKETS = [
  ["Canada", "Alberta & Ontario", "Bookkeeping in QuickBooks & CaseWare, T2 preparation and financial reconstruction for professional corporations, corporations and SMEs."],
  ["Australia", "Healthcare & SMEs", "Bookkeeping in Xero and quarterly BAS preparation for medical practices and SMEs."],
  ["United States", "Tax & advisory", "Tax return preparation and filing, plus tax and business consulting for SMEs."],
  ["Singapore", "Accounting & reporting", "Bookkeeping in QuickBooks & Xero and annual financial statements for the AGM."],
];

const SERVICES = [
  [fa.FaBook, "Bookkeeping & Accounting", "Accurate, reconciled books in QuickBooks, Xero and CaseWare.", ""],
  [fa.FaFileInvoiceDollar, "Tax Return Preparation", "Canadian T2 and Australian BAS, prepared for your firm to review and submit.", "blue"],
  [fa.FaFlagUsa, "US Tax Filing & Consulting", "Tax filing and practical tax and business advice for US SMEs.", ""],
  [fa.FaFileAlt, "Annual Financial Statements", "Prepared for directors' review and the AGM (Singapore).", "blue"],
  [fa.FaTools, "Financial Reconstruction", "Catch-up and clean-up of incomplete records, even years back.", "red"],
];

const WHY = [
  [fa.FaPiggyBank, "Cost Efficiency", "Skilled work at a lower cost than hiring locally."],
  [fa.FaMoon, "Faster Turnaround", "Our time zone keeps work moving while your office is closed."],
  [fa.FaDesktop, "Platform Ready", "Fluent in QuickBooks, Xero, CaseWare and CCH iFirm."],
  [fa.FaChartLine, "Scale on Demand", "Handle busy-season peaks without adding headcount."],
  [fa.FaUserShield, "Confidentiality First", "Client data handled with care and kept private."],
  [fa.FaHandshake, "Your Clients Stay Yours", "We work in the background, under your name and standards."],
];

const brochure = `<!doctype html><html><head><meta charset="utf-8"><title>Lotus BPO Brochure</title>
<style>${CSS}${BROCHURE_CSS}</style></head><body>

<!-- OUTSIDE: inside flap | back | front cover -->
<section class="page">
  <div class="panel w97 tint">
    <div class="eyebrow">Why Lotus BPO</div>
    <h2>Why Firms Partner With Us</h2>
    <div class="stack">
      ${WHY.map(([I, h, p], i) => `<div class="row">${dot(I, i % 2 ? "blue sm" : "sm")}<div><h3>${h}</h3><p>${p}</p></div></div>`).join("")}
    </div>
  </div>

  <div class="panel w100">
    <div class="eyebrow">Getting Started</div>
    <h2>Let's Work Together</h2>
    <div class="steps">
      <div class="step"><div class="num">1</div><div><h3>Talk to Us</h3><p>Tell us about your firm, your clients and the support you need.</p></div></div>
      <div class="step"><div class="num">2</div><div><h3>Start With a Pilot</h3><p>Begin with a few client files and see our quality first-hand.</p></div></div>
      <div class="step"><div class="num">3</div><div><h3>Grow Together</h3><p>Expand the scope as trust and volume grow.</p></div></div>
    </div>
    <div class="serve">
      <div class="eyebrow">Who We Serve</div>
      <div>${dot(fa.FaBriefcase, "sm")} Accounting firms</div>
      <div>${dot(fa.FaStore, "blue sm")} SMEs &amp; corporations</div>
      <div>${dot(fa.FaTooth, "sm")} Professional corporations</div>
      <div>${dot(fa.FaUserMd, "blue sm")} Medical practices</div>
    </div>
    <div class="contact">
      <div class="eyebrow">Contact Us</div>
      <div class="line">${dot(fa.FaPhoneAlt, "red sm")} +84 902 343 218</div>
      <div class="line">${dot(fa.FaMapMarkerAlt, "red sm")} Ho Chi Minh City, Vietnam</div>
    </div>
    <div class="reg">Lotus BPO Company Limited · Enterprise No. 0318715056</div>
  </div>

  <div class="panel w100 dark cover">
    <img class="water" src="../assets/lotus-mark-white.png" style="width:150mm; right:-45mm; bottom:28mm;">
    <div><span class="logo-card"><img src="../assets/logo.png" alt="Lotus BPO"></span></div>
    <h1>Your Trusted Accounting &amp; Tax Partner</h1>
    <div class="sub">Bookkeeping, tax preparation and financial reporting support for accounting firms and growing businesses.</div>
    <div class="markets">
      ${MARKETS.map(([m]) => `<div>${dot(fa.FaMapMarkerAlt, "red")} ${m}</div>`).join("")}
      <div class="tag">One dedicated team in Vietnam</div>
    </div>
  </div>
</section>

<!-- INSIDE: who we are | what we do | where we work -->
<section class="page">
  <div class="panel w100">
    <div class="eyebrow">About Us</div>
    <h2>Who We Are</h2>
    <p class="lead">Lotus BPO is an accounting and tax support company based in Ho Chi Minh City, Vietnam.</p>
    <p class="lead" style="margin-top:3.5mm">We act as an extension of your team, handling bookkeeping, tax preparation and financial clean-ups, so you can focus on advising clients and growing your practice.</p>
    <div class="stats">
      <div class="stat hl"><b>2024</b><span>Founded</span><small>Registered in Ho Chi Minh City</small></div>
      <div class="stat"><b>4</b><span>Markets served</span><small>Canada, Australia, US, Singapore</small></div>
      <div class="stat"><b>4</b><span>Core platforms</span><small>QuickBooks, Xero, CaseWare, CCH iFirm</small></div>
      <div class="stat"><b>6 yrs</b><span>Largest rebuild</span><small>Six years of records, 2021–2026</small></div>
    </div>
    <div class="case">
      <b>Humans in the lead</b>
      <p>Skilled accountants do the work, your team makes the final call, and your clients remain yours.</p>
    </div>
  </div>

  <div class="panel w100" style="border-left:.3mm solid var(--mist); border-right:.3mm solid var(--mist);">
    <div class="eyebrow">Our Services</div>
    <h2>What We Do</h2>
    ${SERVICES.map(([I, h, p, c]) => `<div class="svc">${dot(I, c)}<div><h3>${h}</h3><p>${p}</p></div></div>`).join("")}
    <div class="case">
      <b>Case study: six years rebuilt</b>
      <p>For a dentistry professional corporation in Canada, we rebuilt and reconciled six years of records (2021–2026), ready for the accountant's review and tax filing.</p>
    </div>
  </div>

  <div class="panel w97 tint">
    <div class="eyebrow">Where We Work</div>
    <h2>Four Markets, One Team</h2>
    ${MARKETS.map(([m, t, p]) => `<div class="market"><h3><span class="pin">${ic(fa.FaMapMarkerAlt, "#C8100A")}</span>${m.toUpperCase()}</h3><p><em>${t}.</em> ${p}</p></div>`).join("")}
    <div class="promise"><b>We prepare. You file.</b>When we work through an accounting firm, we prepare the work and the firm reviews and submits. The final sign-off is always yours.</div>
  </div>
</section>
</body></html>`;

const flyer = `<!doctype html><html><head><meta charset="utf-8"><title>Lotus BPO Flyer</title>
<style>${CSS}${FLYER_CSS}</style></head><body>
<section class="flyer">
  <div class="head">
    <img class="water" src="../assets/lotus-mark-white.png" style="width:120mm; right:-18mm; top:8mm;">
    <span class="logo-card"><img src="../assets/logo.png" alt="Lotus BPO"></span>
    <h1>Your Trusted Accounting &amp; Tax Partner</h1>
    <p>An extension of your team in Vietnam: bookkeeping, tax preparation and financial reporting, so you can focus on your clients.</p>
  </div>
  <div class="body">
    <div class="eyebrow">Our Services</div>
    <h2>What We Do</h2>
    <div class="grid">
      ${SERVICES.map(([I, h, p, c]) => `<div class="row">${dot(I, c)}<div><h3>${h}</h3><p>${p}</p></div></div>`).join("")}
      <div class="row">${dot(fa.FaGlobeAsia, "blue")}<div><h3>One Team, Four Markets</h3><p>Canada, Australia, the United States and Singapore.</p></div></div>
    </div>
    <div class="mk">
      ${MARKETS.map(([m, t]) => `<div><h3>${m.toUpperCase()}</h3><p>${t}</p></div>`).join("")}
    </div>
    <div class="why">
      ${WHY.slice(0, 4).map(([, h, p]) => `<div><b>${h}</b><span>${p}</span></div>`).join("")}
    </div>
    <div class="promise">${dot(fa.FaCheckCircle, "red sm")}<div><b>We prepare. You file.</b> When we work through an accounting firm, the firm always reviews and submits. The final sign-off is always yours.</div></div>
  </div>
  <div class="foot">
    <div class="cta"><small>Start with a pilot</small>Let's work together</div>
    <div>
      <div class="line">${dot(fa.FaPhoneAlt, "red sm")} +84 902 343 218</div>
      <div class="line">${dot(fa.FaMapMarkerAlt, "red sm")} Ho Chi Minh City, Vietnam</div>
    </div>
  </div>
</section>
</body></html>`;

(async () => {
  const files = [
    ["brochure.html", brochure, "Lotus BPO - Brochure (A4 Tri-fold).pdf", { width: "297mm", height: "210mm" }],
    ["flyer.html", flyer, "Lotus BPO - Flyer (A4).pdf", { width: "210mm", height: "297mm" }],
  ];
  const browser = await chromium.launch(fs.existsSync("/opt/pw-browsers/chromium") ? {} : {});
  const page = await browser.newPage();
  for (const [html, content, pdf, size] of files) {
    const p = path.join(OUT_DIR, html);
    fs.writeFileSync(p, content);
    await page.goto("file://" + p, { waitUntil: "load" });
    await page.pdf({ path: path.join(__dirname, pdf), ...size, printBackground: true, margin: { top: 0, right: 0, bottom: 0, left: 0 } });
    console.log("Wrote", pdf);
  }
  await browser.close();
})();
