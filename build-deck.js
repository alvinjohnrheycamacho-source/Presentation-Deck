// Builds "Lotus BPO - Conference Deck.pptx".
// Run: node build-deck.js   (needs pptxgenjs, react, react-dom, react-icons, sharp, jszip)
const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const JSZip = require("jszip");
const fa = require("react-icons/fa");

const SKILL = process.env.PPTX_SKILL_DIR;
const OUT = path.join(__dirname, "Lotus BPO - Conference Deck.pptx");
const A = (f) => path.join(__dirname, "assets", f);

// Brand colours sampled from the Lotus BPO logo
const HEX = {
  ink: "1A1F36",
  navy: "1F2A6B",
  red: "C8100A",
  blue: "22428E",
  sky: "5B7BC4",
  mist: "E8ECF7",
  tint: "F3F5FB",
  gray: "5F6577",
};

const THEME = {
  name: "Lotus BPO",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.ink, lt1: "FFFFFF", dk2: HEX.navy, lt2: HEX.tint,
    accent1: HEX.red, accent2: HEX.blue, accent3: HEX.sky,
    accent4: HEX.mist, accent5: HEX.gray, accent6: "0E7C66",
    hlink: HEX.blue, folHlink: HEX.sky,
  },
};

async function icon(Comp, color = "FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + color, size: String(size) })
  );
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
  pres.title = "Lotus BPO Company Limited - Introduction";
  pres.author = "Lotus BPO Company Limited";
  pres.company = "Lotus BPO Company Limited";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const NAVY = C.text2, RED = C.accent1, BLUE = C.accent2, SKY = C.accent3,
    MIST = C.accent4, TINT = C.background2, WHITE = C.background1, INK = C.text1, GRAY = C.accent5;

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "Lotus Dark",
    background: { color: NAVY },
    objects: [
      { image: { path: A("lotus-mark-white.png"), x: 8.3, y: 2.2, w: 6.4, h: 3.95, transparency: 88 } },
      { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.55, w: 8.6, h: 1.7,
        fontSize: 44, bold: true, color: WHITE, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 4.3, w: 8.6, h: 1.0,
        fontSize: 20, color: MIST, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "Lotus Content",
    background: { color: WHITE },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { image: { path: A("logo.png"), x: 11.2, y: 6.92, w: 1.55, h: 0.39 } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 10.4, h: 0.85,
        fontSize: 36, bold: true, color: NAVY, valign: "middle", align: "left", margin: 0 }, text: "" } },
    ],
    slideNumber: { x: 0.6, y: 6.95, w: 0.6, h: 0.35, fontSize: 11, color: GRAY, margin: 0 },
  });

  // helpers
  const T = (s, text, o) => s.addText(text, { isTextBox: true, margin: 0, ...o });
  const circle = (s, name, x, y, d, fill) =>
    s.addShape(pres.shapes.OVAL, { objectName: name, x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
  const card = (s, name, x, y, w, h, fill = TINT) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.12,
      fill: { color: fill }, line: { type: "none" } });
  const shadowCard = (s, name, x, y, w, h) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { objectName: name, x, y, w, h, rectRadius: 0.12,
      fill: { color: WHITE }, line: { color: MIST, width: 1 },
      shadow: { type: "outer", color: "1F2A6B", opacity: 0.12, blur: 8, offset: 2, angle: 90 } });
  const iconDot = async (s, name, Comp, x, y, d, fill, color = "FFFFFF") => {
    circle(s, name, x, y, d, fill);
    s.addImage({ objectName: name + "-i", data: await icon(Comp, color), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 });
  };

  // ---------- 1. Title ----------
  pres.addSection({ title: "Welcome" });
  let s = pres.addSlide({ masterName: "Lotus Dark", sectionTitle: "Welcome" });
  card(s, "a1-logocard", 0.8, 0.75, 3.6, 1.05, WHITE);
  s.addImage({ objectName: "a1-logo", path: A("logo.png"), x: 1.05, y: 0.94, w: 3.1, h: 0.775 });
  s.addText("Your Accounting & Tax Team in Vietnam", { placeholder: "title" });
  s.addText("Introducing Lotus BPO Company Limited", { placeholder: "body" });
  T(s, [
    { text: "Presented to Moores Rowland International & Partner Firms", options: { bold: true, color: SKY, breakLine: true } },
    { text: "MRI Conference & AGM · Mövenpick Resort & Spa, Jimbaran, Bali · 23–25 October 2026", options: { color: MIST, fontSize: 14 } },
  ], { objectName: "a2-for", x: 0.8, y: 5.65, w: 9.5, h: 0.85, fontSize: 16, valign: "top", paraSpaceAfter: 4 });
  s.addNotes("Good day everyone. Thank you to MRI for having us here in Bali. We are Lotus BPO, a company in Ho Chi Minh City, Vietnam. We help accounting firms in other countries with their bookkeeping and tax work. In the next few minutes we will show you who we are, what we do, and how we can work with your firm.");

  // ---------- 2. Who we are ----------
  pres.addSection({ title: "About Us" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "About Us" });
  s.addText("Who We Are", { placeholder: "title" });
  T(s, [
    { text: "Lotus BPO is an accounting support company based in Ho Chi Minh City, Vietnam.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "We work behind the scenes for accounting firms abroad. We do the bookkeeping and draft the tax reports. Your firm reviews and files, and has more time for clients." },
  ], { objectName: "a1-intro", x: 0.6, y: 1.55, w: 5.4, h: 2.3, fontSize: 18, color: INK, valign: "top" });
  // leaders
  T(s, "OUR LEADERS", { objectName: "a2-lh", x: 0.6, y: 4.15, w: 5.4, h: 0.35, fontSize: 12, bold: true, color: RED, charSpacing: 2 });
  const leaders = [
    ["Nguyen Thi Phuong Hien", "General Director"],
    ["Anthony John Galliano", "Director"],
  ];
  for (let i = 0; i < leaders.length; i++) {
    const y = 4.6 + i * 0.85;
    await iconDot(s, "a2-ld" + i, fa.FaUserTie, 0.6, y, 0.65, BLUE);
    T(s, [
      { text: leaders[i][0], options: { bold: true, color: NAVY, fontSize: 16, breakLine: true } },
      { text: leaders[i][1], options: { color: GRAY, fontSize: 14 } },
    ], { objectName: "a2-ldt" + i, x: 1.45, y, w: 4.5, h: 0.65, valign: "middle" });
  }
  // stat tiles 2x2
  const stats = [
    ["2024", "Year we started", "Registered in Vietnam on 15 October 2024"],
    ["2", "Countries we serve", "Canada and Australia"],
    ["3", "Accounting systems", "QuickBooks, CaseWare and Xero"],
    ["6 yrs", "Longest file rebuilt", "A full rebuild, 2021 to 2026"],
  ];
  stats.forEach((st, i) => {
    const x = 6.75 + (i % 2) * 3.05, y = 1.55 + Math.floor(i / 2) * 2.55;
    card(s, `a3-st${i}`, x, y, 2.75, 2.25, i === 0 ? NAVY : TINT);
    const dark = i === 0;
    T(s, st[0], { objectName: `a3-sn${i}`, x: x + 0.3, y: y + 0.25, w: 2.2, h: 0.9, fontSize: 48, bold: true, color: dark ? WHITE : RED });
    T(s, st[1], { objectName: `a3-sl${i}`, x: x + 0.3, y: y + 1.15, w: 2.2, h: 0.4, fontSize: 16, bold: true, color: dark ? WHITE : NAVY });
    T(s, st[2], { objectName: `a3-sd${i}`, x: x + 0.3, y: y + 1.52, w: 2.25, h: 0.6, fontSize: 12, color: dark ? MIST : GRAY, valign: "top" });
  });
  T(s, "Enterprise No. 0318715056 · Ho Chi Minh City Business Registration Office", {
    objectName: "a4-reg", x: 0.6, y: 6.5, w: 9, h: 0.3, fontSize: 10, color: GRAY });
  s.addNotes("Lotus BPO started in October 2024. We are a registered company in Ho Chi Minh City. Today we serve firms in two countries, Canada and Australia. We work every day in QuickBooks, CaseWare and Xero. Our biggest job so far was rebuilding six years of books for one client.");

  // ---------- 3. What we do ----------
  pres.addSection({ title: "Our Services" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Services" });
  s.addText("What We Do", { placeholder: "title" });
  T(s, "Four simple services. One team you can count on.", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 10, h: 0.45, fontSize: 18, color: GRAY });
  const svcs = [
    [fa.FaBook, "Bookkeeping", "We keep your clients' books clean and up to date every month."],
    [fa.FaFileInvoiceDollar, "T2 Return Drafts", "We draft Canadian T2 corporate returns. Your firm reviews and files."],
    [fa.FaCalendarCheck, "BAS Drafts", "We draft quarterly BAS in Xero. Your firm reviews and lodges."],
    [fa.FaTools, "Financial Rebuilds", "Missing or messy records? We rebuild the books, even years back."],
  ];
  for (let i = 0; i < svcs.length; i++) {
    const x = 0.6 + i * 3.08, y = 2.1, w = 2.78;
    shadowCard(s, `a${i + 2}-c`, x, y, w, 3.7);
    await iconDot(s, `a${i + 2}-d`, svcs[i][0], x + 0.35, y + 0.4, 1.0, i % 2 ? BLUE : NAVY);
    T(s, svcs[i][1], { objectName: `a${i + 2}-h`, x: x + 0.35, y: y + 1.65, w: w - 0.6, h: 0.5, fontSize: 20, bold: true, color: NAVY });
    T(s, svcs[i][2], { objectName: `a${i + 2}-t`, x: x + 0.35, y: y + 2.2, w: w - 0.6, h: 1.7, fontSize: 15, color: INK, valign: "top" });
  }
  s.addNotes("Here is what we do. One: bookkeeping. Two: we draft Canadian T2 corporate tax returns. Three: we draft the Australian BAS every quarter. To be clear, we draft; your firm reviews and files. Four: financial rebuilds, when a client's records are missing or behind. We do the heavy lifting, you keep control.");

  // ---------- 4. Where our clients are ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Services" });
  s.addText("Where Our Clients Are", { placeholder: "title" });
  // hub
  circle(s, "a1-hub", 0.9, 2.45, 2.6, NAVY);
  s.addImage({ objectName: "a1-hubi", path: A("lotus-mark-white.png"), x: 1.45, y: 2.95, w: 1.5, h: 0.93 });
  T(s, [{ text: "Our team", options: { breakLine: true, fontSize: 14, color: MIST } }, { text: "Vietnam", options: { fontSize: 22, bold: true, color: WHITE } }],
    { objectName: "a1-hubt", x: 0.9, y: 3.95, w: 2.6, h: 0.8, align: "center" });
  // arrows
  s.addShape(pres.shapes.LINE, { objectName: "a2-l1", x: 3.6, y: 2.75, w: 1.5, h: 0.75, flipV: true, line: { color: SKY, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { objectName: "a2-l2", x: 3.6, y: 4.2, w: 1.5, h: 0.75, line: { color: SKY, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  const regions = [
    ["CANADA", "Alberta & Ontario", "Bookkeeping, T2 return drafts and financial rebuilds", "Professional Corporations and Corporations", fa.FaMapMarkerAlt],
    ["AUSTRALIA", "Medical field", "Bookkeeping in Xero and quarterly BAS drafts", "Doctors and medical practices", fa.FaMapMarkerAlt],
  ];
  for (let i = 0; i < 2; i++) {
    const y = 1.55 + i * 2.65, x = 5.3;
    card(s, `a${3 + i}-c`, x, y, 7.4, 2.35, TINT);
    await iconDot(s, `a${3 + i}-d`, regions[i][4], x + 0.35, y + 0.35, 0.75, RED);
    T(s, regions[i][0], { objectName: `a${3 + i}-h`, x: x + 1.3, y: y + 0.35, w: 3, h: 0.45, fontSize: 24, bold: true, color: NAVY, charSpacing: 1 });
    T(s, regions[i][1], { objectName: `a${3 + i}-s`, x: x + 1.3, y: y + 0.8, w: 5.8, h: 0.35, fontSize: 15, bold: true, color: RED });
    T(s, [
      { text: "What we do: ", options: { bold: true, color: NAVY } }, { text: regions[i][2], options: { breakLine: true } },
      { text: "Who we serve: ", options: { bold: true, color: NAVY } }, { text: regions[i][3] },
    ], { objectName: `a${3 + i}-t`, x: x + 1.3, y: y + 1.25, w: 5.8, h: 0.9, fontSize: 15, color: INK, paraSpaceAfter: 4, valign: "top" });
  }
  s.addNotes("Our team is in Vietnam. Our clients are in Canada, in Alberta and Ontario, and in Australia. In Canada we work with Professional Corporations and regular corporations. In Australia our clients are doctors and medical practices.");

  // ---------- 5. Canada ----------
  pres.addSection({ title: "Canada" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Canada" });
  s.addText("Canada: Our Work", { placeholder: "title" });
  const ca = [
    [fa.FaBook, "Bookkeeping", "In QuickBooks and CaseWare"],
    [fa.FaFileInvoiceDollar, "T2 Return Drafts", "We draft the T2. Your firm reviews and files it."],
    [fa.FaTools, "Financial Rebuilds", "When records are missing or years behind"],
  ];
  for (let i = 0; i < ca.length; i++) {
    const y = 1.6 + i * 1.5;
    await iconDot(s, `a${i + 1}-d`, ca[i][0], 0.6, y, 0.95, i === 2 ? RED : NAVY);
    T(s, [
      { text: ca[i][1], options: { bold: true, fontSize: 20, color: NAVY, breakLine: true } },
      { text: ca[i][2], options: { fontSize: 15, color: GRAY } },
    ], { objectName: `a${i + 1}-t`, x: 1.8, y: y - 0.05, w: 5.0, h: 1.05, valign: "middle" });
  }
  // right panel
  card(s, "a4-p", 7.2, 1.55, 5.5, 4.95, NAVY);
  T(s, "WHO WE SERVE", { objectName: "a4-h", x: 7.6, y: 1.9, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
  T(s, [
    { text: "Professional Corporations", options: { bullet: true, breakLine: true } },
    { text: "Corporations (legal companies)", options: { bullet: true } },
  ], { objectName: "a4-t", x: 7.6, y: 2.3, w: 4.8, h: 0.95, fontSize: 17, color: WHITE, paraSpaceAfter: 6, valign: "top" });
  T(s, "OUR CANADIAN PARTNERS", { objectName: "a5-h", x: 7.6, y: 3.3, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
  const partners = [
    ["Opty Business Solutions Inc.", "We draft T2s in their CCH iFirm. Opty files."],
    ["Harry Chang, CA", "We draft T2s in an Excel workbook. They file in their own system."],
  ];
  partners.forEach((p, i) => {
    const y = 3.72 + i * 1.08;
    card(s, `a5-pc${i}`, 7.6, y, 4.7, 0.95, WHITE);
    T(s, [
      { text: p[0], options: { bold: true, fontSize: 16, color: NAVY, breakLine: true } },
      { text: p[1], options: { fontSize: 13, color: GRAY } },
    ], { objectName: `a5-pt${i}`, x: 7.85, y, w: 4.3, h: 0.95, valign: "middle" });
  });
  T(s, "New clients keep coming through these firms.", { objectName: "a5-n", x: 7.6, y: 5.95, w: 4.8, h: 0.35, fontSize: 14, italic: true, color: MIST });
  s.addNotes("In Canada, we do bookkeeping in QuickBooks and CaseWare. We draft the T2 corporate tax returns each year, but we do not file them. With Opty, we draft the T2 inside their CCH iFirm system and Opty files. With Harry Chang, we prepare an Excel T2 workbook and they file it in their own system. We also rebuild financial records when needed. We work through two Canadian partners: Opty Business Solutions and Harry Chang, Chartered Accountant. New clients come to us through them all the time, both Professional Corporations and regular corporations.");

  // ---------- 6. Case story ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Canada" });
  s.addText("Case Story: 6 Years Rebuilt", { placeholder: "title" });
  T(s, "Client: A Dentistry Professional Corporation in Canada", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 10, h: 0.45, fontSize: 18, color: GRAY });
  // timeline
  const yrs = [2021, 2022, 2023, 2024, 2025, 2026];
  s.addShape(pres.shapes.LINE, { objectName: "a2-tl", x: 1.2, y: 2.55, w: 10.9, h: 0, line: { color: MIST, width: 4 } });
  yrs.forEach((yr, i) => {
    const x = 0.85 + i * 2.18;
    circle(s, `a${2}-y${i}`, x + 0.17, 2.2, 0.7, i === yrs.length - 1 ? RED : NAVY);
    T(s, "✓", { objectName: `a2-yc${i}`, x: x + 0.17, y: 2.2, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, String(yr), { objectName: `a2-yt${i}`, x: x - 0.2, y: 3.0, w: 1.45, h: 0.4, fontSize: 18, bold: true, color: NAVY, align: "center" });
  });
  const story = [
    ["The Need", "Six years of books, from 2021 to 2026, had to be rebuilt from the start."],
    ["What We Did", "We gathered the records, rebuilt the books year by year and matched every account."],
    ["The Result", "Complete, clean records for every year, ready for the accountant and the tax filings."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    card(s, `a${3 + i}-c`, x, 3.85, 3.8, 2.3, i === 2 ? NAVY : TINT);
    T(s, story[i][0], { objectName: `a${3 + i}-h`, x: x + 0.3, y: 4.1, w: 3.2, h: 0.45, fontSize: 20, bold: true, color: i === 2 ? WHITE : RED });
    T(s, story[i][1], { objectName: `a${3 + i}-t`, x: x + 0.3, y: 4.6, w: 3.2, h: 1.7, fontSize: 15, color: i === 2 ? WHITE : INK, valign: "top" });
  }
  s.addNotes("Here is a recent example. A dentistry Professional Corporation in Canada needed six years of books rebuilt, from 2021 to 2026. We collected the records, rebuilt each year, and matched every account. The client now has clean, complete records, ready for the accountant and for tax filing. If your firm has clients who are behind, we can help.");

  // ---------- 7. Australia ----------
  pres.addSection({ title: "Australia" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Australia" });
  s.addText("Australia: Our Work", { placeholder: "title" });
  T(s, "Partner firm: Stone Group Accounting Firm  ·  Our clients: doctors and medical practices", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 12, h: 0.45, fontSize: 18, color: GRAY });
  // flow: Xero Blue -> Xero Red -> ATO
  const flow = [
    [fa.FaBook, "Xero Blue", "The client's books", "We do the bookkeeping and draft the BAS every quarter.", BLUE, "LOTUS BPO DOES THIS"],
    [fa.FaCalculator, "Xero Red", "The firm's tax side", "The figures flow across automatically. The firm reviews the BAS.", RED, "YOUR FIRM DOES THIS"],
    [fa.FaLandmark, "ATO", "Australian Taxation Office", "The firm lodges the BAS. We never lodge on our own.", NAVY, ""],
  ];
  for (let i = 0; i < flow.length; i++) {
    const x = 0.6 + i * 4.15, y = 2.35, w = 3.65, g = `a${i + 2}`;
    if (flow[i][5]) T(s, flow[i][5], { objectName: g + "-tag", x, y: 1.9, w: 6, h: 0.3, fontSize: 12, bold: true, color: i === 0 ? BLUE : RED, charSpacing: 2 });
    card(s, g + "-c", x, y, w, 3.45, flow[i][4]);
    await iconDot(s, g + "-d", flow[i][0], x + 0.3, y + 0.3, 0.85, WHITE, i === 0 ? HEX.blue : i === 1 ? HEX.red : HEX.navy);
    T(s, flow[i][1], { objectName: g + "-h", x: x + 0.3, y: y + 1.3, w: w - 0.6, h: 0.5, fontSize: 24, bold: true, color: WHITE });
    T(s, flow[i][2], { objectName: g + "-s", x: x + 0.3, y: y + 1.8, w: w - 0.6, h: 0.35, fontSize: 14, bold: true, color: MIST });
    T(s, flow[i][3], { objectName: g + "-t", x: x + 0.3, y: y + 2.2, w: w - 0.5, h: 1.1, fontSize: 15, color: WHITE, valign: "top" });
    if (i < flow.length - 1)
      s.addShape(pres.shapes.RIGHT_ARROW, { objectName: `a${i + 3}-ar`, x: x + w + 0.08, y: y + 1.5, w: 0.34, h: 0.45, fill: { color: SKY }, line: { type: "none" } });
  }
  T(s, [
    { text: "Simply put: ", options: { bold: true, color: NAVY } },
    { text: "Xero Blue is where the client's books live. Xero Red is the accounting firm's side of Xero, linked to the tax office." },
  ], { objectName: "a5-n", x: 0.6, y: 6.05, w: 10.4, h: 0.6, fontSize: 14, color: INK, valign: "middle" });
  s.addNotes("In Australia we work with Stone Group Accounting Firm, and our clients are doctors and medical practices. We work in Xero. Xero Blue is the client's own books: that is where we do the bookkeeping and draft the BAS every quarter. Xero Red is the accounting firm's side of Xero, which is linked to the tax office. The figures flow across, the firm reviews the BAS, and the firm lodges it with the ATO. We draft; we never lodge on our own.");

  // ---------- 8. We draft, you file ----------
  pres.addSection({ title: "Working Together" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("We Draft. You File.", { placeholder: "title" });
  T(s, "How it works today with each of our partner firms", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 10, h: 0.45, fontSize: 18, color: GRAY });
  const cols = [[0.6, 3.6, "PARTNER FIRM"], [4.4, 4.1, "WHAT LOTUS BPO DRAFTS"], [8.7, 4.0, "WHO FILES"]];
  cols.forEach((c, i) => T(s, c[2], { objectName: "a2-ch" + i, x: c[0] + 0.25, y: 1.9, w: c[1], h: 0.3, fontSize: 12, bold: true, color: i === 2 ? RED : BLUE, charSpacing: 2 }));
  const rows = [
    ["Opty Business Solutions Inc.", "Canada", "T2 return, prepared in their CCH iFirm", "Opty reviews and files"],
    ["Harry Chang, CA", "Canada", "T2 workbook, prepared in Excel", "Harry Chang files in their own system"],
    ["Stone Group Accounting Firm", "Australia", "Quarterly BAS, prepared in Xero", "Stone Group reviews and lodges"],
  ];
  for (let i = 0; i < rows.length; i++) {
    const y = 2.3 + i * 1.08, g = `a${i + 3}`;
    card(s, g + "-r", 0.6, y, 12.1, 0.92, i % 2 ? WHITE : TINT);
    T(s, [
      { text: rows[i][0], options: { bold: true, fontSize: 16, color: NAVY, breakLine: true } },
      { text: rows[i][1], options: { fontSize: 13, color: GRAY } },
    ], { objectName: g + "-p", x: 0.85, y, w: 3.5, h: 0.92, valign: "middle" });
    T(s, rows[i][2], { objectName: g + "-w", x: 4.65, y, w: 3.95, h: 0.92, fontSize: 15, color: INK, valign: "middle" });
    T(s, rows[i][3], { objectName: g + "-f", x: 8.95, y, w: 3.6, h: 0.92, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
  }
  card(s, "a6-k", 0.6, 5.6, 12.1, 0.95, NAVY);
  await iconDot(s, "a6-d", fa.FaCheckCircle, 0.85, 5.72, 0.7, RED);
  T(s, [
    { text: "To be clear: ", options: { bold: true, color: WHITE } },
    { text: "we draft the tax reports, but we never file or lodge them ourselves. Your firm always checks and submits.", options: { color: MIST } },
  ], { objectName: "a6-t", x: 1.75, y: 5.6, w: 10.8, h: 0.95, fontSize: 16, valign: "middle" });
  s.addNotes("We want to be open about this. We draft the tax reports, but we do not file them. With Opty, we draft the T2 in their CCH iFirm and Opty files it. With Harry Chang, we prepare an Excel T2 workbook and they file it in their own system. With Stone Group in Australia, we draft the BAS in Xero and Stone Group lodges it. Your firm always has the final say.");

  // ---------- 9. How we work ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("How We Work With Your Firm", { placeholder: "title" });
  const steps = [
    [fa.FaComments, "Talk", "We learn what you and your clients need."],
    [fa.FaKey, "Set Up", "You give us access to the files and software."],
    [fa.FaLaptop, "We Do the Work", "Our team does the books and drafts the tax reports."],
    [fa.FaSearch, "You Review", "Your team checks and signs off."],
    [fa.FaPaperPlane, "You File", "You file from your own system and keep the client."],
  ];
  for (let i = 0; i < steps.length; i++) {
    const x = 0.6 + i * 2.48;
    const g = `a${i + 1}`;
    if (i < steps.length - 1)
      s.addShape(pres.shapes.LINE, { objectName: g + "-ln", x: x + 1.45, y: 2.25, w: 1.5, h: 0, line: { color: SKY, width: 2, dashType: "dash" } });
    await iconDot(s, g + "-d", steps[i][0], x + 0.5, 1.65, 1.2, i === 2 ? RED : NAVY);
    T(s, `STEP ${i + 1}`, { objectName: g + "-n", x, y: 3.0, w: 2.2, h: 0.3, fontSize: 12, bold: true, color: RED, align: "center", charSpacing: 2 });
    T(s, steps[i][1], { objectName: g + "-h", x, y: 3.3, w: 2.2, h: 0.5, fontSize: 19, bold: true, color: NAVY, align: "center" });
    T(s, steps[i][2], { objectName: g + "-t", x: x + 0.05, y: 3.85, w: 2.1, h: 1.0, fontSize: 14, color: INK, align: "center", valign: "top" });
  }
  card(s, "a6-k", 0.6, 5.35, 12.1, 1.15, TINT);
  await iconDot(s, "a6-d", fa.FaHandshake, 0.9, 5.52, 0.8, RED);
  T(s, [
    { text: "Humans in the lead. ", options: { bold: true, color: NAVY } },
    { text: "Real accountants do the work, your team makes the final call, and your clients stay your clients." },
  ], { objectName: "a6-t", x: 1.95, y: 5.35, w: 10.5, h: 1.15, fontSize: 17, color: INK, valign: "middle" });
  s.addNotes("Working with us is simple. First, we talk and learn your needs. Second, you give us access to the files and software. Third, we do the books and draft the tax reports. Fourth, your team reviews and signs off. Fifth, you file from your own system. This fits the theme of this conference, Humans in the Lead: real people do the work, your team makes the final call, and your clients stay your clients.");

  // ---------- 9. Why us ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("Why Partner With Lotus BPO", { placeholder: "title" });
  const why = [
    [fa.FaPiggyBank, "Lower Costs", "Get quality work at a lower cost than hiring locally."],
    [fa.FaMoon, "Work Done Overnight", "Vietnam's time zone means work moves while you sleep."],
    [fa.FaDesktop, "We Know Your Software", "QuickBooks, CaseWare and Xero, used every day."],
    [fa.FaChartLine, "Grow Without Hiring", "Take on more clients in busy season with no new staff."],
    [fa.FaUserShield, "Care With Your Data", "Client files are handled with care and kept private."],
    [fa.FaHandshake, "Proven Partnerships", "Trusted today by firms in Canada and Australia."],
  ];
  for (let i = 0; i < why.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.15, y = 1.75 + row * 1.75;
    await iconDot(s, `a${row + 1}-d${col}`, why[i][0], x, y, 0.95, col ? BLUE : NAVY);
    T(s, [
      { text: why[i][1], options: { bold: true, fontSize: 20, color: NAVY, breakLine: true } },
      { text: why[i][2], options: { fontSize: 15, color: INK } },
    ], { objectName: `a${row + 1}-t${col}`, x: x + 1.2, y: y - 0.1, w: 4.65, h: 1.2, valign: "middle" });
  }
  s.addNotes("Why work with us? Lower costs. Work done overnight thanks to the time difference. We already use your software every day. You can grow in busy season without hiring. We take care with client data. And firms in Canada and Australia already trust us.");

  // ---------- 10. New office ----------
  pres.addSection({ title: "Our Office" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Office" });
  s.addText("Our New Office", { placeholder: "title" });
  T(s, "Opened on August 10, 2026 in Ho Chi Minh City, with room for our growing team.", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 12, h: 0.45, fontSize: 18, color: GRAY });
  // photos 4:3
  const pw = 5.9, ph = pw * 0.75;
  s.addImage({ objectName: "a2-p1", path: A("office-1.jpg"), x: 0.6, y: 1.95, w: pw, h: ph * 0.95, sizing: { type: "cover", w: pw, h: ph * 0.95 }, rounding: false });
  s.addImage({ objectName: "a3-p2", path: A("office-2.jpg"), x: 6.8, y: 1.95, w: pw, h: ph * 0.95, sizing: { type: "cover", w: pw, h: ph * 0.95 } });
  s.addNotes("This is our new office in Ho Chi Minh City. We signed and moved in on August 10, 2026. It gives our team more space to grow as we take on more clients.");

  // ---------- 11. Let's work together ----------
  pres.addSection({ title: "Next Steps" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Next Steps" });
  s.addText("Let's Work Together", { placeholder: "title" });
  const next = [
    [fa.FaCoffee, "Meet Us in Bali", "Find us at the coffee breaks. We would love to hear about your firm."],
    [fa.FaFolderOpen, "Start Small", "Begin with a few client files and see how we work."],
    [fa.FaSeedling, "Grow Together", "Add more clients and services as trust builds."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    shadowCard(s, `a${i + 1}-c`, x, 1.6, 3.8, 3.15);
    T(s, String(i + 1), { objectName: `a${i + 1}-n`, x: x + 2.6, y: 1.75, w: 1.0, h: 0.9, fontSize: 48, bold: true, color: MIST, align: "right" });
    await iconDot(s, `a${i + 1}-d`, next[i][0], x + 0.35, 1.9, 0.9, i === 2 ? RED : NAVY);
    T(s, next[i][1], { objectName: `a${i + 1}-h`, x: x + 0.35, y: 3.0, w: 3.1, h: 0.45, fontSize: 21, bold: true, color: NAVY });
    T(s, next[i][2], { objectName: `a${i + 1}-t`, x: x + 0.35, y: 3.5, w: 3.15, h: 1.1, fontSize: 15, color: INK, valign: "top" });
  }
  card(s, "a4-k", 0.6, 5.1, 12.1, 1.4, NAVY);
  await iconDot(s, "a4-d1", fa.FaPhoneAlt, 0.95, 5.4, 0.8, RED);
  T(s, [
    { text: "Call us", options: { fontSize: 13, color: SKY, bold: true, breakLine: true } },
    { text: "+84 902 343 218", options: { fontSize: 22, bold: true, color: WHITE } },
  ], { objectName: "a4-t1", x: 1.95, y: 5.25, w: 4.0, h: 1.1, valign: "middle" });
  await iconDot(s, "a4-d2", fa.FaMapMarkerAlt, 6.6, 5.4, 0.8, RED);
  T(s, [
    { text: "Visit us", options: { fontSize: 13, color: SKY, bold: true, breakLine: true } },
    { text: "Ho Chi Minh City, Vietnam", options: { fontSize: 22, bold: true, color: WHITE } },
  ], { objectName: "a4-t2", x: 7.6, y: 5.25, w: 4.9, h: 1.1, valign: "middle" });
  s.addNotes("So, what's next? Come and meet us here in Bali, at any coffee break. Start small with a few client files and see how we work. Then grow with us as trust builds. You can call us anytime at +84 902 343 218.");

  // ---------- 12. Thank you ----------
  s = pres.addSlide({ masterName: "Lotus Dark", sectionTitle: "Next Steps" });
  card(s, "a1-logocard", 0.8, 0.75, 3.6, 1.05, WHITE);
  s.addImage({ objectName: "a1-logo", path: A("logo.png"), x: 1.05, y: 0.94, w: 3.1, h: 0.775 });
  s.addText("Thank You", { placeholder: "title" });
  s.addText("Lotus BPO Company Limited · Your accounting & tax team in Vietnam", { placeholder: "body" });
  T(s, "Questions? We are happy to talk.", { objectName: "a2-q", x: 0.8, y: 5.85, w: 8.6, h: 0.45, fontSize: 16, color: SKY, bold: true });
  s.addNotes("Thank you for your time. We are happy to answer any questions, here or after the session.");

  await pres.writeFile({ fileName: OUT });
  if (SKILL) {
    const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
    await applyTheme(OUT, THEME);
  }
  await addMotion(OUT);
  console.log("Wrote", OUT);
})();

// ---------- Transitions + entrance animations ----------
// Shapes named "aN-..." fade in automatically, group by group (N = order).
async function addMotion(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const f of slides) {
    let xml = await zip.file(f).async("string");
    const groups = {};
    const re = /<p:(sp|pic|cxnSp)>[\s\S]*?<p:cNvPr id="(\d+)" name="a(\d+)-[^"]*"/g;
    // match each shape's own cNvPr (non-greedy across shapes is guarded by splitting)
    const parts = xml.split(/(?=<p:(?:sp|pic|cxnSp)>)/);
    for (const p of parts) {
      const m = /^<p:(sp|pic|cxnSp)>[\s\S]*?<p:cNvPr id="(\d+)" name="a(\d+)-/.exec(p);
      if (m && p.indexOf("<p:cNvPr") === p.search(/<p:cNvPr/)) {
        (groups[m[3]] = groups[m[3]] || []).push({ id: m[2], kind: m[1], text: m[1] === "sp" });
      }
    }
    let id = 3, timing = "";
    const order = Object.keys(groups).map(Number).sort((a, b) => a - b);
    if (order.length) {
      const STEP = 450, DUR = 500;
      let pars = "";
      order.forEach((g, gi) => {
        let eff = "";
        groups[g].forEach((sh, si) => {
          const a = ++id, b = ++id, c = ++id;
          eff += `<p:par><p:cTn id="${a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="${si === 0 ? (gi === 0 ? "afterEffect" : "afterEffect") : "withEffect"}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>` +
            `<p:set><p:cBhvr><p:cTn id="${b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${sh.id}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>` +
            `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${c}" dur="${DUR}"/><p:tgtEl><p:spTgt spid="${sh.id}"/></p:tgtEl></p:cBhvr></p:animEffect>` +
            `</p:childTnLst></p:cTn></p:par>`;
        });
        const pid = ++id;
        pars += `<p:par><p:cTn id="${pid}" fill="hold"><p:stCondLst><p:cond delay="${gi === 0 ? 300 : 300 + gi * STEP}"/></p:stCondLst><p:childTnLst>${eff}</p:childTnLst></p:cTn></p:par>`;
      });
      const bld = order.flatMap((g) => groups[g]).filter((s) => s.text)
        .map((s) => `<p:bldP spid="${s.id}" grpId="0" animBg="1"/>`).join("");
      timing = `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>` +
        `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
        `<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>` +
        pars +
        `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>` +
        `</p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ""}</p:timing>`;
    }
    const n = Number(/slide(\d+)\.xml/.exec(f)[1]);
    const trans = n === 1 ? `<p:transition spd="slow"><p:fade/></p:transition>` : `<p:transition spd="med"><p:push dir="l"/></p:transition>`;
    xml = xml.replace(/<p:transition[\s\S]*?<\/p:transition>|<p:transition[^>]*\/>/g, "").replace(/<p:timing>[\s\S]*?<\/p:timing>/g, "");
    if (xml.includes("</p:clrMapOvr>")) xml = xml.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + trans + timing);
    else xml = xml.replace("</p:cSld>", "</p:cSld>" + trans + timing);
    zip.file(f, xml);
  }
  fs.writeFileSync(file, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}
