'use strict';
/**
 * Generates output/CarrierBooking_SupplierGuide.pptx
 * Simple, layman-friendly 6-slide deck for TEST & REACT SUPPLIERS explaining what the
 * Carrier Booking Tool needs from them, what happens after they hit "send", the
 * validation checks we run, and how we communicate with them (incl. contact email).
 * Slide 4 includes a compact technical flow summary for IT/EDI contacts.
 * Run: node generate-supplier-guide-pptx.js
 */
const PptxGenJS = require('pptxgenjs');
const path = require('path');
const fs   = require('fs');

const OUT = path.join(__dirname, 'output');
if (!fs.existsSync(OUT)) fs.mkdirSync(OUT, { recursive: true });

const pptx = new PptxGenJS();
pptx.layout = 'LAYOUT_WIDE';

const C = {
  navy:    '0B1E3A', blue:    '1565C0', accent:  '0288D1', cyan:    '38BDF8',
  mint:    '059669', amber:   'D97706', red:     'DC2626', pink:    'DB2777',
  white:   'FFFFFF', surface: 'F1F5F9', muted:   '64748B', dark:    '0F172A',
  lightBg: 'EFF6FF', greenBg: 'ECFDF5', amberBg: 'FFFBEB',
  asosBlack: '0E0F0F', asosGrey: '6B6B6B',
};

function contentSlide(titleText, subText) {
  const s = pptx.addSlide();
  s.background = { color: C.surface };
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:1.0, fill:{color:C.navy}, line:{color:C.navy} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:1.0, w:13.33, h:0.06, fill:{color:C.accent}, line:{color:C.accent} });
  s.addText(titleText, { x:0.35, y:0, w:12.6, h:1.0, fontSize:22, bold:true, color:C.white, fontFace:'Calibri', align:'left', valign:'middle' });
  if (subText) s.addText(subText, { x:0.35, y:1.12, w:12.6, h:0.32, fontSize:12.5, italic:true, color:C.muted, fontFace:'Calibri' });
  return s;
}

function box(slide, text, x, y, w, h, bg, fg, fsize) {
  slide.addShape(pptx.ShapeType.roundRect, { x,y,w,h, fill:{color:bg}, line:{color:bg}, rectRadius:0.08 });
  slide.addText(text, { x,y,w,h, fontSize:fsize||13, bold:true, color:fg||C.white, fontFace:'Calibri', align:'center', valign:'middle', wrap:true });
}

// ---- SLIDE 1: Title ----------------------------------------------------------
{
  const s = pptx.addSlide();
  s.background = { color: C.asosBlack };
  s.addText('Carrier Booking Made Simple', {
    x:0.72, y:1.6, w:11.9, h:1.4,
    fontSize:42, bold:true, color:C.white, fontFace:'Calibri Light', align:'left', valign:'top', wrap:true
  });
  s.addText('A Quick Guide for Our Suppliers — What We Need From You & What Happens Next', {
    x:0.72, y:3.05, w:11.5, h:0.6,
    fontSize:18, color:C.white, fontFace:'Calibri Light', align:'left', valign:'middle'
  });
  s.addShape(pptx.ShapeType.roundRect, {x:0.72,y:4.0,w:11.5,h:0.7, fill:{color:'1A2A1F'}, line:{color:C.mint,pt:1.5}, rectRadius:0.08});
  s.addText('No AIM login. No XML. Just fill one simple Excel sheet and email it to us — we do the rest.', {
    x:0.9,y:4.0,w:11.2,h:0.7, fontSize:14, color:'A7F3D0', fontFace:'Calibri', valign:'middle'
  });
  s.addText('ASOS — Test & React Inbound Team  |  Supplier Guide', {
    x:0.72, y:6.9, w:9.18, h:0.45,
    fontSize:11, color:C.asosGrey, fontFace:'Calibri', align:'left', valign:'middle'
  });
}

// ---- SLIDE 2: Why we're asking + the 3 steps (merged) ------------------------
{
  const s = contentSlide('Why We\u2019re Asking For This — It\u2019s Just 3 Steps');

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:1.15,w:12.62,h:1.55, fill:{color:C.greenBg}, line:{color:C.mint,pt:1.5}, rectRadius:0.1});
  s.addText([
    {text:'Carrier bookings used to need AIM — a system built for ASOS staff, not suppliers. ', options:{fontSize:13.5, color:C.dark}},
    {text:'Now you just fill one simple Excel sheet and email it to us — ', options:{bold:true, fontSize:13.5, color:C.dark}},
    {text:'no login, no XML, no new system. We match it to your PO and book it with the carrier automatically.', options:{fontSize:13.5, color:C.dark}},
  ], {x:0.5,y:1.15,w:12.32,h:1.55, fontFace:'Calibri', valign:'middle', wrap:true});

  const steps = [
    {n:'1', title:'Fill the Template', desc:'Enter your PO number, booking\ndate, carton count and weight.', color:C.blue},
    {n:'2', title:'Email It To Us',    desc:'Attach the completed sheet and\nsend it the same way you already\nsend other documents.', color:C.accent},
    {n:'3', title:'We Take It From Here', desc:'We check it, match it to your\nPO, and book it with the carrier\nautomatically.', color:C.mint},
  ];
  const bw=3.9, bh=3.9, by=3.05, gap=0.36;
  steps.forEach((st,i) => {
    const bx = 0.4 + i*(bw+gap);
    s.addShape(pptx.ShapeType.ellipse, {x:bx+bw/2-0.4, y:by, w:0.8,h:0.8, fill:{color:st.color}, line:{color:C.white,pt:2}});
    s.addText(st.n, {x:bx+bw/2-0.4, y:by, w:0.8,h:0.8, fontSize:26, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:by+0.95,w:bw,h:2.85, fill:{color:C.white}, line:{color:st.color,pt:2}, rectRadius:0.12});
    s.addText(st.title, {x:bx+0.15,y:by+1.1,w:bw-0.3,h:0.55, fontSize:16, bold:true, color:st.color, fontFace:'Calibri', align:'center'});
    s.addText(st.desc,  {x:bx+0.25,y:by+1.7,w:bw-0.5,h:2.0, fontSize:13, color:C.dark, fontFace:'Calibri', align:'center', valign:'top', wrap:true});
    if (i < steps.length-1)
      s.addShape(pptx.ShapeType.rightArrow, {x:bx+bw+0.02, y:by+2.25, w:gap-0.05, h:0.3, fill:{color:C.muted}, line:{color:C.muted}});
  });
}

// ---- SLIDE 3: What we need from you — the template ---------------------------
{
  const s = contentSlide('What We Need From You', 'One Excel sheet, filled in per PO — coloured guide tells you exactly what to complete');

  const sections = [
    { y:1.55, bg:C.pink, label:'MUST FILL (pink) — required for every PO',
      body:'PO Number   |   Cargo Ready / Collection Date   |   Booking Request Date   |   Traffic Mode (CFS or CY)   |   Carton Type', bh:0.42 },
    { y:2.45, bg:'1B7F3E', label:'PRE-FILLED FOR YOU (green) — check it, change only if different',
      body:'Booking Group   |   Pack Type   |   Collection Type   |   Hazardous Goods flag', bh:0.42 },
    { y:3.35, bg:C.accent, label:'CALCULATED AUTOMATICALLY (blue) — don\u2019t touch these',
      body:'Carton length / width / height / weight — filled in from the Carton Type you picked', bh:0.42 },
    { y:4.25, bg:C.muted, label:'OPTIONAL (grey) — only if it applies to your shipment',
      body:'Collection Time   |   Remarks / Comments', bh:0.42 },
  ];
  sections.forEach(sec => {
    s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:sec.y, w:12.62,h:0.42, fill:{color:sec.bg}, line:{color:sec.bg}, rectRadius:0.06});
    s.addText(sec.label, {x:0.45,y:sec.y, w:12.42,h:0.42, fontSize:13.5, bold:true, color:C.white, fontFace:'Calibri', valign:'middle'});
    s.addText(sec.body,  {x:0.5, y:sec.y+0.43, w:12.4, h:0.4, fontSize:12.5, color:C.dark, fontFace:'Calibri', valign:'top', wrap:true});
  });

  // Sample template snippet as a real table
  s.addText('Example — one filled-in row from the template:', {x:0.35,y:5.15,w:8,h:0.3, fontSize:12, bold:true, color:C.dark, fontFace:'Calibri'});
  const headerOpts = { bold:true, color:'FFFFFF', fill:{color:C.navy}, fontSize:9.5, align:'center', valign:'middle' };
  const dataOpts   = { color:C.dark, fill:{color:'FFFFFF'}, fontSize:9.5, align:'center', valign:'middle' };
  const rows = [
    [
      {text:'PO Number', options:headerOpts}, {text:'Collection Date', options:headerOpts},
      {text:'Booking Req. Date', options:headerOpts}, {text:'Traffic Mode', options:headerOpts},
      {text:'Carton Type', options:headerOpts}, {text:'No. of Cartons', options:headerOpts},
      {text:'Unit Weight (KG)', options:headerOpts},
    ],
    [
      {text:'500036995371', options:dataOpts}, {text:'23/09/2026', options:dataOpts},
      {text:'24/09/2026', options:dataOpts}, {text:'CFS', options:dataOpts},
      {text:'BDCM1', options:dataOpts}, {text:'48', options:dataOpts}, {text:'0.29', options:dataOpts},
    ],
  ];
  s.addTable(rows, { x:0.35, y:5.48, w:12.62, h:0.7, colW:[1.9,1.75,1.95,1.6,1.5,1.9,2.02], border:{type:'solid', color:'CBD5E1', pt:0.75} });
}

// ---- SLIDE 4: What happens after you hit send --------------------------------
{
  const s = contentSlide('What Happens After You Hit \u201cSend\u201d');

  const steps = [
    {n:'1', label:'You email the\nfilled template\nto ASOS',            color:C.blue  },
    {n:'2', label:'It lands in\nour inbox\nautomatically',              color:C.accent},
    {n:'3', label:'We check it against\nyour PO in our\nsystems',        color:C.mint  },
    {n:'4', label:'A carrier booking\nis created\nautomatically',        color:C.amber },
    {n:'5', label:'Sent straight to\nthe carrier\n(Davis Turner)',       color:C.pink  },
    {n:'6', label:'Carrier plans\nyour collection',                     color:C.navy  },
  ];
  const bw=1.9, bh=2.0, by=1.35, gap=0.18;
  steps.forEach((st,i) => {
    const bx = 0.35 + i*(bw+gap);
    s.addShape(pptx.ShapeType.ellipse, {x:bx+bw/2-0.28, y:by-0.38, w:0.56,h:0.56, fill:{color:st.color}, line:{color:st.color}});
    s.addText(st.n, {x:bx+bw/2-0.28, y:by-0.38, w:0.56,h:0.56, fontSize:16, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:by,w:bw,h:bh, fill:{color:st.color}, line:{color:st.color}, rectRadius:0.1});
    s.addText(st.label, {x:bx,y:by,w:bw,h:bh, fontSize:12.5, color:C.white, fontFace:'Calibri', align:'center', valign:'middle', wrap:true});
    if (i < steps.length-1)
      s.addShape(pptx.ShapeType.rightArrow, {x:bx+bw, y:by+bh/2-0.14, w:gap, h:0.28, fill:{color:C.muted}, line:{color:C.muted}});
  });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:3.75,w:12.62,h:0.6, fill:{color:C.greenBg}, line:{color:C.mint}, rectRadius:0.07});
  s.addText('This all happens automatically, usually within a few hours of you sending the email — no need to chase or follow up.', {
    x:0.5,y:3.75,w:12.3,h:0.6, fontSize:13.5, color:C.mint, fontFace:'Calibri', bold:true, align:'left', valign:'middle', wrap:true
  });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:4.5,w:12.62,h:0.6, fill:{color:C.amberBg}, line:{color:C.amber}, rectRadius:0.07});
  s.addText('If something doesn\u2019t match (see next slide), we\u2019ll flag it and reach out to you before the booking is sent.', {
    x:0.5,y:4.5,w:12.3,h:0.6, fontSize:13.5, color:C.amber, fontFace:'Calibri', align:'left', valign:'middle', wrap:true
  });

  // one-page technical detail strip for IT/EDI contacts
  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:5.28,w:12.62,h:1.5, fill:{color:C.navy}, line:{color:C.navy}, rectRadius:0.08});
  s.addText('For your IT / EDI team — the technical flow in one line:', {x:0.5,y:5.36,w:12.3,h:0.3, fontSize:11.5, bold:true, color:C.cyan, fontFace:'Calibri'});
  s.addText(
    'Supplier email (.xlsx attachment)  \u2192  Power Automate (Outlook trigger)  \u2192  Azure Blob / SharePoint  \u2192  Azure App Service ' +
    '(Node.js)  \u2192  Template parsed & validated  \u2192  Matched to PO/ASN in ASOS Data Estate (Databricks)  \u2192  VBKREQ XML generated  \u2192  ' +
    'Uploaded to E2open via SFTP  \u2192  Davis Turner (carrier) receives booking.',
    {x:0.5,y:5.66,w:12.3,h:1.0, fontSize:12, color:C.white, fontFace:'Calibri', valign:'top', wrap:true}
  );
}

// ---- SLIDE 5: Checks we run + How we communicate with you --------------------
{
  const s = contentSlide('Checks We Run \u2014 And How We Stay In Touch');

  const checks = [
    { title:'Required fields filled in?', desc:'PO, dates, cartons and weight must all be present, or we flag it back to you.', color:C.red },
    { title:'Quantities match our records?', desc:'We compare your totals against the PO/ASN already in ASOS\u2019 systems and flag any mismatch.', color:C.amber },
    { title:'Still an active PO/shipment?', desc:'Already cancelled or already booked? We skip it automatically \u2014 no duplicates.', color:C.accent },
    { title:'Values sensible?', desc:'Whole-number units; Collection Time required if Collection Type = "Collection".', color:C.mint },
  ];
  const bw=6.1, bh=1.55, gapX=0.42, gapY=0.18;
  checks.forEach((c,i) => {
    const col = i % 2, row = Math.floor(i/2);
    const bx = 0.35 + col*(bw+gapX);
    const bY = 1.2 + row*(bh+gapY);
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:bY,w:bw,h:bh, fill:{color:C.white}, line:{color:c.color,pt:2}, rectRadius:0.08});
    s.addShape(pptx.ShapeType.rect, {x:bx,y:bY,w:0.1,h:bh, fill:{color:c.color}, line:{color:c.color}});
    s.addText(c.title, {x:bx+0.28,y:bY+0.1,w:bw-0.45,h:0.4, fontSize:13, bold:true, color:c.color, fontFace:'Calibri', valign:'top', wrap:true});
    s.addText(c.desc,  {x:bx+0.28,y:bY+0.5,w:bw-0.45,h:0.95, fontSize:11, color:C.dark, fontFace:'Calibri', valign:'top', wrap:true});
  });

  // Communication section
  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:4.35,w:12.62,h:2.35, fill:{color:C.lightBg}, line:{color:C.blue,pt:1.5}, rectRadius:0.1});
  s.addText('How We Communicate With You', {x:0.5,y:4.45,w:12.3,h:0.4, fontSize:15, bold:true, color:C.blue, fontFace:'Calibri'});
  s.addText([
    {text:'You send: ', options:{bold:true, fontSize:12.5, color:C.dark}},
    {text:'one email with the completed template to InboundService@asos.com.\n', options:{fontSize:12.5, color:C.dark}},
    {text:'We confirm: ', options:{bold:true, fontSize:12.5, color:C.dark}},
    {text:'if anything is missing or doesn\u2019t match, we email you back before the booking is finalised.\n', options:{fontSize:12.5, color:C.dark}},
    {text:'Status checks: ', options:{bold:true, fontSize:12.5, color:C.dark}},
    {text:'email us any time for an update on a specific PO or booking.\n', options:{fontSize:12.5, color:C.dark}},
    {text:'Internal reporting: ', options:{bold:true, fontSize:12.5, color:C.dark}},
    {text:'our team receives an automatic booking status report twice a day (09:00 / 13:00) to track everything end-to-end.', options:{fontSize:12.5, color:C.dark}},
  ], {x:0.55,y:4.9,w:12.3,h:1.7, fontFace:'Calibri', valign:'top', wrap:true, lineSpacing:19});
}

// ---- SLIDE 6: Checklist + Contact / Thank you --------------------------------
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:1.0, fill:{color:C.navy}, line:{color:C.navy} });
  s.addText('Quick Checklist \u2014 Then We\u2019re Here to Help', { x:0.35, y:0, w:12.6, h:1.0, fontSize:24, bold:true, color:C.white, fontFace:'Calibri', align:'left', valign:'middle' });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:1.2,w:6.0,h:4.4, fill:{color:'0F2A1B'}, line:{color:C.mint,pt:2}, rectRadius:0.1});
  s.addText('DO', {x:0.5,y:1.3,w:5.7,h:0.45, fontSize:16, bold:true, color:C.mint, fontFace:'Calibri'});
  s.addText(
    '\u2713  Use the latest template we sent you\n' +
    '\u2713  Fill in every pink (mandatory) field\n' +
    '\u2713  Double-check the PO number matches exactly\n' +
    '\u2713  Enter dates as DD/MM/YYYY\n' +
    '\u2713  Send one email per shipment/PO batch',
    {x:0.55,y:1.8,w:5.6,h:3.7, fontSize:13, color:C.white, fontFace:'Calibri', valign:'top', wrap:true, lineSpacing:24}
  );

  s.addShape(pptx.ShapeType.roundRect, {x:6.7,y:1.2,w:6.3,h:4.4, fill:{color:'2A1414'}, line:{color:C.red,pt:2}, rectRadius:0.1});
  s.addText('DON\u2019T', {x:6.85,y:1.3,w:6.0,h:0.45, fontSize:16, bold:true, color:C.red, fontFace:'Calibri'});
  s.addText(
    '\u2717  Don\u2019t edit the blue (auto-filled) carton fields\n' +
    '\u2717  Don\u2019t rename the template columns\n' +
    '\u2717  Don\u2019t send a screenshot or PDF instead of Excel\n' +
    '\u2717  Don\u2019t worry about mistakes \u2014 resend a corrected file and let us know',
    {x:6.85,y:1.8,w:6.0,h:3.7, fontSize:13, color:C.white, fontFace:'Calibri', valign:'top', wrap:true, lineSpacing:24}
  );

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:5.85,w:12.62,h:1.05, fill:{color:C.accent}, line:{color:C.accent}, rectRadius:0.1});
  s.addText('Questions any time: InboundService@asos.com', {
    x:0.5,y:5.85,w:12.3,h:1.05, fontSize:20, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'
  });
}

const outFile = path.join(OUT, 'CarrierBooking_SupplierGuide.pptx');
pptx.writeFile({ fileName: outFile }).then(() => {
  console.log('Saved: ' + outFile);
}).catch(err => { console.error('Error:', err.message); process.exit(1); });

