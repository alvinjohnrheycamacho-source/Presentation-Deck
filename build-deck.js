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
  s.addText("Your Trusted Accounting & Tax Partner", { placeholder: "title" });
  s.addText("Introducing Lotus BPO Company Limited · Ho Chi Minh City, Vietnam", { placeholder: "body" });
  T(s, [
    { text: "Presented to Moores Rowland International & Partner Firms", options: { bold: true, color: SKY, breakLine: true } },
    { text: "MRI Conference & AGM · Mövenpick Resort & Spa, Jimbaran, Bali · 23–25 October 2026", options: { color: MIST, fontSize: 14 } },
  ], { objectName: "a2-for", x: 0.8, y: 5.65, w: 9.5, h: 0.85, fontSize: 16, valign: "top", paraSpaceAfter: 4 });
  s.addNotes("Good day, everyone, and thank you to MRI for having us here in Bali. We are Lotus BPO, an accounting and tax support company based in Ho Chi Minh City, Vietnam. Over the next few minutes, I will share who we are, the markets we serve, and how we can support your firm and your clients.");

  // ---------- 2. Who we are ----------
  pres.addSection({ title: "About Us" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "About Us" });
  s.addText("Who We Are", { placeholder: "title" });
  T(s, [
    { text: "Lotus BPO is an accounting and tax support company based in Ho Chi Minh City, Vietnam.", options: { breakLine: true } },
    { text: " ", options: { breakLine: true, fontSize: 8 } },
    { text: "We act as an extension of your team, handling bookkeeping, tax preparation and financial clean-ups, so you can focus on advising clients and growing your practice." },
  ], { objectName: "a1-intro", x: 0.6, y: 1.55, w: 5.4, h: 2.4, fontSize: 18, color: INK, valign: "top" });
  T(s, "OUR LEADERSHIP", { objectName: "a2-lh", x: 0.6, y: 4.15, w: 5.4, h: 0.35, fontSize: 12, bold: true, color: RED, charSpacing: 2 });
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
  const stats = [
    ["2024", "Founded", "Registered in Ho Chi Minh City, Vietnam"],
    ["4", "Markets served", "Canada, Australia, the US and Singapore"],
    ["4", "Core platforms", "QuickBooks, Xero, CaseWare and CCH iFirm"],
    ["6 yrs", "Largest rebuild", "Six years of records rebuilt, 2021–2026"],
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
  s.addNotes("Lotus BPO was founded in 2024 and is registered in Ho Chi Minh City. We see ourselves as an extension of our clients' teams. We take care of the bookkeeping, tax preparation and clean-up work, so they can spend more time advising clients. Today we serve four markets: Canada, Australia, the United States and Singapore. Our team works every day in QuickBooks, Xero, CaseWare and CCH iFirm, and our largest project so far was rebuilding six years of financial records for a single client.");

  // ---------- 3. What we do ----------
  pres.addSection({ title: "Our Services" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Services" });
  s.addText("What We Do", { placeholder: "title" });
  T(s, "End-to-end accounting and tax support, tailored to each market.", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 11, h: 0.45, fontSize: 18, color: GRAY });
  const svcs = [
    [fa.FaBook, "Bookkeeping & Accounting", "Accurate books in QuickBooks, Xero and CaseWare, plus year-end financial statements."],
    [fa.FaFileInvoiceDollar, "Tax Return Preparation", "Canadian T2 and Australian BAS, prepared for your firm to review and submit."],
    [fa.FaFlagUsa, "US Tax Filing", "Tax preparation and filing for small and mid-sized US businesses."],
    [fa.FaLightbulb, "Tax & Business Consulting", "Practical guidance that helps US SMEs make sound decisions."],
    [fa.FaTools, "Financial Reconstruction", "Catch-up and clean-up of incomplete records, even several years back."],
  ];
  for (let i = 0; i < 6; i++) {
    const x = 0.6 + (i % 3) * 4.125, y = 1.95 + Math.floor(i / 3) * 2.45, w = 3.85, h = 2.2, g = `a${Math.floor(i / 3) + 2}`;
    if (i === 5) {
      card(s, g + "-c5", x, y, w, h, NAVY);
      T(s, [
        { text: "4 markets.", options: { fontSize: 28, bold: true, color: WHITE, breakLine: true } },
        { text: "1 dedicated team.", options: { fontSize: 28, bold: true, color: SKY, breakLine: true } },
        { text: "Canada · Australia · US · Singapore", options: { fontSize: 14, color: MIST } },
      ], { objectName: g + "-t5", x: x + 0.35, y, w: w - 0.6, h, valign: "middle", paraSpaceAfter: 4 });
      break;
    }
    shadowCard(s, `${g}-c${i}`, x, y, w, h);
    await iconDot(s, `${g}-d${i}`, svcs[i][0], x + 0.3, y + 0.3, 0.8, i % 2 ? BLUE : NAVY);
    T(s, svcs[i][1], { objectName: `${g}-h${i}`, x: x + 1.25, y: y + 0.3, w: w - 1.45, h: 0.8, fontSize: 18, bold: true, color: NAVY, valign: "middle" });
    T(s, svcs[i][2], { objectName: `${g}-t${i}`, x: x + 0.3, y: y + 1.25, w: w - 0.55, h: 0.85, fontSize: 15, color: INK, valign: "top" });
  }
  s.addNotes("Here is how we help. We keep accurate books in QuickBooks, Xero and CaseWare, and prepare year-end financial statements. We prepare Canadian T2 returns and Australian BAS for our partner firms to review and submit. For US small and mid-sized businesses, we handle tax filing and provide tax and business consulting. And when records are incomplete or behind, we rebuild them, even several years back. Four markets, one dedicated team.");

  // ---------- 4. Global reach ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Services" });
  s.addText("Our Global Reach", { placeholder: "title" });
  circle(s, "a1-hub", 0.7, 2.6, 2.6, NAVY);
  s.addImage({ objectName: "a1-hubi", path: A("lotus-mark-white.png"), x: 1.25, y: 3.1, w: 1.5, h: 0.93 });
  T(s, [{ text: "One team in", options: { breakLine: true, fontSize: 14, color: MIST } }, { text: "Vietnam", options: { fontSize: 22, bold: true, color: WHITE } }],
    { objectName: "a1-hubt", x: 0.7, y: 4.1, w: 2.6, h: 0.8, align: "center" });
  s.addShape(pres.shapes.LINE, { objectName: "a2-l1", x: 3.4, y: 2.75, w: 1.3, h: 0.9, flipV: true, line: { color: SKY, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  s.addShape(pres.shapes.LINE, { objectName: "a2-l2", x: 3.4, y: 4.25, w: 1.3, h: 0.9, line: { color: SKY, width: 2.5, dashType: "dash", endArrowType: "triangle" } });
  const regions = [
    ["CANADA", "Alberta & Ontario", "Bookkeeping, T2 preparation, financial reconstruction", "Professional corporations, corporations & SMEs"],
    ["AUSTRALIA", "Healthcare & SMEs", "Xero bookkeeping, quarterly BAS preparation", "Medical practices & SMEs"],
    ["UNITED STATES", "Tax & advisory", "Tax filing, tax & business consulting", "Small & mid-sized businesses"],
    ["SINGAPORE", "Accounting & reporting", "Bookkeeping in QuickBooks & Xero, annual financial statements for the AGM", "Small & mid-sized businesses"],
  ];
  for (let i = 0; i < 4; i++) {
    const x = 4.9 + (i % 2) * 3.97, y = 1.55 + Math.floor(i / 2) * 2.55, w = 3.75, g = `a${3 + i}`;
    card(s, g + "-c", x, y, w, 2.3, TINT);
    await iconDot(s, g + "-d", fa.FaMapMarkerAlt, x + 0.25, y + 0.25, 0.6, RED);
    T(s, regions[i][0], { objectName: g + "-h", x: x + 1.0, y: y + 0.22, w: w - 1.1, h: 0.4, fontSize: 19, bold: true, color: NAVY, charSpacing: 1 });
    T(s, regions[i][1], { objectName: g + "-s", x: x + 1.0, y: y + 0.6, w: w - 1.1, h: 0.3, fontSize: 14, bold: true, color: RED });
    T(s, [
      { text: "Services: ", options: { bold: true, color: NAVY } }, { text: regions[i][2], options: { breakLine: true } },
      { text: "Clients: ", options: { bold: true, color: NAVY } }, { text: regions[i][3] },
    ], { objectName: g + "-t", x: x + 0.25, y: y + 1.05, w: w - 0.45, h: 1.15, fontSize: 14, color: INK, paraSpaceAfter: 4, valign: "top" });
  }
  s.addNotes("From one team in Vietnam, we now serve four markets. In Canada, mainly Alberta and Ontario, we support professional corporations, corporations and SMEs. In Australia, we work with medical practices and SMEs. In the United States, we provide tax filing and consulting for small and mid-sized businesses. And in Singapore, we handle bookkeeping in QuickBooks and Xero and prepare annual financial statements for the AGM.");

  // ---------- 5. Who we serve ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Our Services" });
  s.addText("Who We Serve", { placeholder: "title" });
  T(s, "From accounting firms to growing businesses, our support fits the way you work.", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 12, h: 0.45, fontSize: 18, color: GRAY });
  const who = [
    [fa.FaBriefcase, "Accounting Firms", "We work as your back office, under your name and your quality standards."],
    [fa.FaStore, "SMEs & Corporations", "Small and mid-sized businesses and incorporated companies in all four markets."],
    [fa.FaTooth, "Professional Corporations", "Dental and other professional corporations across Canada."],
    [fa.FaUserMd, "Medical Practices", "Doctors and clinics in Australia, supported every quarter."],
  ];
  for (let i = 0; i < who.length; i++) {
    const x = 0.6 + i * 3.08, y = 2.05, w = 2.78;
    shadowCard(s, `a${i + 2}-c`, x, y, w, 3.6);
    await iconDot(s, `a${i + 2}-d`, who[i][0], x + 0.35, y + 0.4, 1.0, i % 2 ? BLUE : NAVY);
    T(s, who[i][1], { objectName: `a${i + 2}-h`, x: x + 0.35, y: y + 1.6, w: w - 0.55, h: 0.75, fontSize: 19, bold: true, color: NAVY, valign: "top" });
    T(s, who[i][2], { objectName: `a${i + 2}-t`, x: x + 0.35, y: y + 2.35, w: w - 0.6, h: 1.4, fontSize: 15, color: INK, valign: "top" });
  }
  s.addNotes("Who do we serve? First, accounting firms, for whom we act as a back office under their own name and standards. Second, small and mid-sized businesses and corporations in all four markets. Third, professional corporations in Canada, including dental practices. And fourth, doctors and clinics in Australia.");

  // ---------- 6. Canada ----------
  pres.addSection({ title: "Canada" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Canada" });
  s.addText("Canada", { placeholder: "title" });
  const ca = [
    [fa.FaBook, "Bookkeeping & Accounting", "In QuickBooks and CaseWare"],
    [fa.FaFileInvoiceDollar, "T2 Return Preparation", "We prepare the T2; your firm reviews and files it"],
    [fa.FaTools, "Financial Reconstruction", "Bringing incomplete or overdue records up to date"],
  ];
  for (let i = 0; i < ca.length; i++) {
    const y = 1.6 + i * 1.5;
    await iconDot(s, `a${i + 1}-d`, ca[i][0], 0.6, y, 0.95, i === 2 ? RED : NAVY);
    T(s, [
      { text: ca[i][1], options: { bold: true, fontSize: 20, color: NAVY, breakLine: true } },
      { text: ca[i][2], options: { fontSize: 15, color: GRAY } },
    ], { objectName: `a${i + 1}-t`, x: 1.8, y: y - 0.05, w: 5.1, h: 1.05, valign: "middle" });
  }
  card(s, "a4-p", 7.2, 1.55, 5.5, 4.95, NAVY);
  T(s, "WHO WE SERVE", { objectName: "a4-h", x: 7.6, y: 1.85, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
  T(s, [
    { text: "Professional corporations", options: { bullet: true, breakLine: true } },
    { text: "Incorporated companies", options: { bullet: true, breakLine: true } },
    { text: "Small & mid-sized enterprises (SMEs)", options: { bullet: true } },
  ], { objectName: "a4-t", x: 7.6, y: 2.22, w: 4.8, h: 1.15, fontSize: 16, color: WHITE, paraSpaceAfter: 4, valign: "top" });
  T(s, "OUR CANADIAN PARTNERS", { objectName: "a5-h", x: 7.6, y: 3.5, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
  const partners = [
    ["Opty Business Solutions Inc.", "We prepare T2s in their CCH iFirm; Opty files."],
    ["Harry Chang, CA", "We prepare an Excel T2 workbook; they file in their own system."],
  ];
  partners.forEach((p, i) => {
    const y = 3.9 + i * 1.03;
    card(s, `a5-pc${i}`, 7.6, y, 4.7, 0.92, WHITE);
    T(s, [
      { text: p[0], options: { bold: true, fontSize: 16, color: NAVY, breakLine: true } },
      { text: p[1], options: { fontSize: 13, color: GRAY } },
    ], { objectName: `a5-pt${i}`, x: 7.85, y, w: 4.3, h: 0.92, valign: "middle" });
  });
  T(s, "A steady flow of new clients through both firms.", { objectName: "a5-n", x: 7.6, y: 5.98, w: 4.8, h: 0.35, fontSize: 14, italic: true, color: MIST });
  s.addNotes("In Canada, we keep the books in QuickBooks and CaseWare, prepare T2 corporate returns, and rebuild records that are incomplete or overdue. Our clients include professional corporations, incorporated companies and SMEs. We work through two partner firms. With Opty Business Solutions, we prepare the T2 inside their CCH iFirm and Opty files it. With Harry Chang, Chartered Accountant, we prepare an Excel T2 workbook and they file it in their own system. Both firms bring us a steady flow of new clients.");

  // ---------- 7. Case story ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Canada" });
  s.addText("Case Study: Six Years Rebuilt", { placeholder: "title" });
  T(s, "Client: A Dentistry Professional Corporation in Canada", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 10, h: 0.45, fontSize: 18, color: GRAY });
  const yrs = [2021, 2022, 2023, 2024, 2025, 2026];
  s.addShape(pres.shapes.LINE, { objectName: "a2-tl", x: 1.2, y: 2.55, w: 10.9, h: 0, line: { color: MIST, width: 4 } });
  yrs.forEach((yr, i) => {
    const x = 0.85 + i * 2.18;
    circle(s, `a2-y${i}`, x + 0.17, 2.2, 0.7, i === yrs.length - 1 ? RED : NAVY);
    T(s, "✓", { objectName: `a2-yc${i}`, x: x + 0.17, y: 2.2, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: WHITE, align: "center", valign: "middle" });
    T(s, String(yr), { objectName: `a2-yt${i}`, x: x - 0.2, y: 3.0, w: 1.45, h: 0.4, fontSize: 18, bold: true, color: NAVY, align: "center" });
  });
  const story = [
    ["The Challenge", "Six years of financial records, from 2021 to 2026, had to be rebuilt from the ground up."],
    ["Our Approach", "We gathered the source documents, rebuilt the books year by year and reconciled every account."],
    ["The Outcome", "Complete, reliable records for every year, ready for the accountant's review and tax filing."],
  ];
  for (let i = 0; i < 3; i++) {
    const x = 0.6 + i * 4.1;
    card(s, `a${3 + i}-c`, x, 3.85, 3.8, 2.3, i === 2 ? NAVY : TINT);
    T(s, story[i][0], { objectName: `a${3 + i}-h`, x: x + 0.3, y: 4.1, w: 3.2, h: 0.45, fontSize: 20, bold: true, color: i === 2 ? WHITE : RED });
    T(s, story[i][1], { objectName: `a${3 + i}-t`, x: x + 0.3, y: 4.6, w: 3.2, h: 1.45, fontSize: 15, color: i === 2 ? WHITE : INK, valign: "top" });
  }
  s.addNotes("Here is one recent project. A dentistry professional corporation in Canada needed six years of records, from 2021 to 2026, rebuilt from the ground up. We gathered the source documents, rebuilt the books year by year and reconciled every account. The result: complete, reliable records for every year, ready for the accountant's review and tax filing. If your firm has clients who have fallen behind, this is exactly where we can help.");

  // ---------- 8. Australia ----------
  pres.addSection({ title: "Australia" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Australia" });
  s.addText("Australia", { placeholder: "title" });
  T(s, "Partner firm: Stone Group Accounting Firm  ·  Clients: medical practices and SMEs", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 12, h: 0.45, fontSize: 18, color: GRAY });
  const flow = [
    [fa.FaBook, "Xero Blue", "The client's books", "We keep the books and prepare the BAS every quarter.", BLUE, "LOTUS BPO"],
    [fa.FaCalculator, "Xero Red", "The firm's tax side", "Figures flow across automatically. The firm reviews the BAS.", RED, "YOUR FIRM"],
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
    { text: "In short: ", options: { bold: true, color: NAVY } },
    { text: "Xero Blue is where the client's books live. Xero Red is the accounting firm's side of Xero, linked to the ATO." },
  ], { objectName: "a5-n", x: 0.6, y: 6.05, w: 10.4, h: 0.6, fontSize: 14, color: INK, valign: "middle" });
  s.addNotes("In Australia, we work with Stone Group Accounting Firm, supporting medical practices and SMEs. Everything runs in Xero. Xero Blue is the client's own books: this is where we keep the bookkeeping current and prepare the BAS every quarter. Xero Red is the accounting firm's side of Xero, which is linked to the ATO. The figures flow across automatically, the firm reviews the BAS, and the firm lodges it. We prepare; the firm always lodges.");

  // ---------- 9-10. United States, Singapore ----------
  // Market slide in the Canada style: services on the left, navy panel on the right.
  const marketSlide = async (section, title, sub, services, serve, listTitle, list, notes) => {
    pres.addSection({ title: section });
    s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: section });
    s.addText(title, { placeholder: "title" });
    T(s, sub, { objectName: "a1-sub", x: 0.6, y: 1.25, w: 6.4, h: 0.45, fontSize: 18, color: GRAY });
    for (let i = 0; i < services.length; i++) {
      const y = 2.05 + i * 1.45;
      await iconDot(s, `a${i + 2}-d`, services[i][0], 0.6, y, 0.95, i === services.length - 1 ? RED : NAVY);
      T(s, [
        { text: services[i][1], options: { bold: true, fontSize: 20, color: NAVY, breakLine: true } },
        { text: services[i][2], options: { fontSize: 15, color: GRAY } },
      ], { objectName: `a${i + 2}-t`, x: 1.8, y: y - 0.05, w: 5.1, h: 1.05, valign: "middle" });
    }
    card(s, "a5-p", 7.2, 1.55, 5.5, 3.25 + (list.length - 1) * 0.8 + 0.66 + 0.3 - 1.55, NAVY);
    T(s, "WHO WE SERVE", { objectName: "a5-h", x: 7.6, y: 1.85, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
    T(s, serve, { objectName: "a5-t", x: 7.6, y: 2.2, w: 4.8, h: 0.4, fontSize: 17, bold: true, color: WHITE });
    T(s, listTitle, { objectName: "a6-h", x: 7.6, y: 2.85, w: 4.8, h: 0.35, fontSize: 13, bold: true, color: SKY, charSpacing: 2 });
    list.forEach((item, i) => {
      const y = 3.25 + i * 0.8;
      card(s, `a6-c${i}`, 7.6, y, 4.7, 0.66, WHITE);
      T(s, item, { objectName: `a6-n${i}`, x: 7.85, y, w: 4.3, h: 0.66, fontSize: 15, bold: true, color: NAVY, valign: "middle" });
    });
    s.addNotes(notes);
  };

  await marketSlide("United States", "United States",
    "Tax and advisory support for small and mid-sized businesses",
    [
      [fa.FaFileInvoiceDollar, "Tax Return Preparation & Filing", "Prepared accurately and filed on schedule"],
      [fa.FaLightbulb, "Tax & Business Consulting", "Practical guidance for everyday decisions"],
      [fa.FaChartLine, "Support That Scales", "Our help grows as the business grows"],
    ],
    "Small & mid-sized businesses (SMEs)",
    "WHAT OUR CLIENTS GET",
    ["One dedicated team for their tax needs", "Clear answers in plain language", "Deadlines tracked and met"],
    "In the United States, we work with small and mid-sized businesses. We prepare and file their tax returns, and we provide tax and business consulting to help owners make sound decisions. Our clients get one dedicated team, clear answers in plain language, and deadlines that are tracked and met.");

  await marketSlide("Singapore", "Singapore",
    "Accounting and reporting for small and mid-sized businesses",
    [
      [fa.FaBook, "Bookkeeping in QuickBooks & Xero", "Books kept accurate and up to date"],
      [fa.FaBalanceScale, "Accounting & Records", "Clean, reconciled accounts all year round"],
      [fa.FaFileAlt, "Annual Financial Statements", "Prepared for directors' review and the AGM"],
    ],
    "Small & mid-sized businesses (SMEs)",
    "SELECTED CLIENTS",
    ["Samai Distillery Cambodia Pte., Ltd.", "Nurkas Holdings Pte. Ltd.", "Qwqer International Pte. Ltd.", "Tanncam Investment Pte. Ltd."],
    "In Singapore, we support small and mid-sized businesses. We keep their books in QuickBooks and Xero, maintain clean, reconciled accounts throughout the year, and prepare their annual financial statements, ready for the directors' review and the company's Annual General Meeting. Our clients include Samai Distillery Cambodia, Nurkas Holdings, Qwqer International and Tanncam Investment.");

  // ---------- 10. We prepare, you file ----------
  pres.addSection({ title: "Working Together" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("We Prepare. You File.", { placeholder: "title" });
  T(s, "How we work alongside our accounting-firm partners today", {
    objectName: "a1-sub", x: 0.6, y: 1.25, w: 10, h: 0.45, fontSize: 18, color: GRAY });
  const cols = [[0.6, 3.6, "PARTNER FIRM"], [4.4, 4.1, "WHAT LOTUS BPO PREPARES"], [8.7, 4.0, "WHO FILES"]];
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
    { text: "Our commitment: ", options: { bold: true, color: WHITE } },
    { text: "when we work through an accounting firm, we prepare and the firm reviews and submits. The final sign-off is always yours.", options: { color: MIST } },
  ], { objectName: "a6-t", x: 1.75, y: 5.6, w: 10.8, h: 0.95, fontSize: 16, valign: "middle" });
  s.addNotes("We want to be fully transparent about our role. When we work through an accounting firm, we prepare the tax reports and the firm reviews and submits them. With Opty, we prepare the T2 in their CCH iFirm and Opty files it. With Harry Chang, we prepare an Excel T2 workbook and they file it in their own system. With Stone Group, we prepare the BAS in Xero and Stone Group lodges it. The final sign-off always stays with the firm.");

  // ---------- 11. How we work ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("How We Work With Your Firm", { placeholder: "title" });
  const steps = [
    [fa.FaComments, "Discover", "We learn your needs, systems and deadlines."],
    [fa.FaKey, "Onboard", "Secure access to your files and software is set up."],
    [fa.FaLaptop, "Prepare", "Our team completes the books and prepares the tax reports."],
    [fa.FaSearch, "Review", "Your team checks and approves the work."],
    [fa.FaPaperPlane, "Submit", "You file from your own system and keep the client."],
  ];
  for (let i = 0; i < steps.length; i++) {
    const x = 0.6 + i * 2.48, g = `a${i + 1}`;
    if (i < steps.length - 1)
      s.addShape(pres.shapes.LINE, { objectName: g + "-ln", x: x + 1.45, y: 2.25, w: 1.5, h: 0, line: { color: SKY, width: 2, dashType: "dash" } });
    await iconDot(s, g + "-d", steps[i][0], x + 0.5, 1.65, 1.2, i === 2 ? RED : NAVY);
    T(s, `STEP ${i + 1}`, { objectName: g + "-n", x, y: 3.0, w: 2.2, h: 0.3, fontSize: 12, bold: true, color: RED, align: "center", charSpacing: 2 });
    T(s, steps[i][1], { objectName: g + "-h", x, y: 3.3, w: 2.2, h: 0.5, fontSize: 19, bold: true, color: NAVY, align: "center" });
    T(s, steps[i][2], { objectName: g + "-t", x: x + 0.05, y: 3.85, w: 2.1, h: 1.1, fontSize: 14, color: INK, align: "center", valign: "top" });
  }
  card(s, "a6-k", 0.6, 5.35, 12.1, 1.15, TINT);
  await iconDot(s, "a6-d", fa.FaHandshake, 0.9, 5.52, 0.8, RED);
  T(s, [
    { text: "Humans in the lead. ", options: { bold: true, color: NAVY } },
    { text: "Skilled accountants do the work, your team makes the final call, and your clients remain yours." },
  ], { objectName: "a6-t", x: 1.95, y: 5.35, w: 10.5, h: 1.15, fontSize: 17, color: INK, valign: "middle" });
  s.addNotes("Working with us follows five simple steps. Discover: we learn your needs, systems and deadlines. Onboard: we set up secure access to your files and software. Prepare: our team completes the books and prepares the tax reports. Review: your team checks and approves. Submit: you file from your own system and keep the client relationship. This reflects the theme of this conference, Humans in the Lead: skilled accountants do the work, your team makes the final call, and your clients remain yours.");

  // ---------- 12. Why us ----------
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Working Together" });
  s.addText("Why Partner With Lotus BPO", { placeholder: "title" });
  const why = [
    [fa.FaPiggyBank, "Cost Efficiency", "Skilled work at a lower cost than hiring locally."],
    [fa.FaMoon, "Faster Turnaround", "Our time zone keeps work moving while your office is closed."],
    [fa.FaDesktop, "Platform Ready", "Fluent in QuickBooks, Xero, CaseWare and CCH iFirm."],
    [fa.FaChartLine, "Scale on Demand", "Handle busy-season peaks without adding headcount."],
    [fa.FaUserShield, "Confidentiality First", "Client data handled with care and kept private."],
    [fa.FaGlobeAsia, "Proven in Four Markets", "Trusted in Canada, Australia, the US and Singapore."],
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
  s.addNotes("Why partner with us? Cost efficiency: skilled work at a lower cost than hiring locally. Faster turnaround: thanks to the time difference, work keeps moving while your office is closed. Platform ready: we already work in QuickBooks, Xero, CaseWare and CCH iFirm. Scale on demand: handle busy season without adding headcount. Confidentiality first: your client data is handled with care. And we are proven in four markets.");

  // ---------- 13. Let's work together ----------
  pres.addSection({ title: "Next Steps" });
  s = pres.addSlide({ masterName: "Lotus Content", sectionTitle: "Next Steps" });
  s.addText("Let's Work Together", { placeholder: "title" });
  const next = [
    [fa.FaCoffee, "Meet Us in Bali", "Let's talk over coffee during the conference breaks."],
    [fa.FaFolderOpen, "Start With a Pilot", "Begin with a few client files and see our quality first-hand."],
    [fa.FaSeedling, "Grow Together", "Expand the scope as trust and volume grow."],
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
  s.addNotes("So, where do we go from here? Let's talk over coffee during the conference breaks. If it feels like a good fit, start with a pilot of a few client files and see our quality first-hand. Then we grow the scope together as trust and volume build. You can reach us anytime at +84 902 343 218.");

  // ---------- 14. Thank you ----------
  s = pres.addSlide({ masterName: "Lotus Dark", sectionTitle: "Next Steps" });
  card(s, "a1-logocard", 0.8, 0.75, 3.6, 1.05, WHITE);
  s.addImage({ objectName: "a1-logo", path: A("logo.png"), x: 1.05, y: 0.94, w: 3.1, h: 0.775 });
  s.addText("Thank You", { placeholder: "title" });
  s.addText("Lotus BPO Company Limited · Your trusted accounting & tax partner", { placeholder: "body" });
  T(s, "Questions? We would be glad to hear from you.", { objectName: "a2-q", x: 0.8, y: 5.85, w: 8.6, h: 0.45, fontSize: 16, color: SKY, bold: true });
  s.addNotes("Thank you for your time and attention. I would be glad to take any questions, here or during the breaks.");

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
