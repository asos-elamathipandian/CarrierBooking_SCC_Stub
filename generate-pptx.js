'use strict';
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
  lightBg: 'EFF6FF', greenBg: 'ECFDF5',
  asosBlack: '0E0F0F', asosGrey: '6B6B6B',
};

function contentSlide(titleText) {
  const s = pptx.addSlide();
  s.background = { color: C.surface };
  s.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:1.0, fill:{color:C.navy}, line:{color:C.navy} });
  s.addShape(pptx.ShapeType.rect, { x:0, y:1.0, w:13.33, h:0.06, fill:{color:C.accent}, line:{color:C.accent} });
  s.addText(titleText, { x:0.35, y:0, w:12.6, h:1.0, fontSize:22, bold:true, color:C.white, fontFace:'Calibri', align:'left', valign:'middle' });
  return s;
}

function box(slide, text, x, y, w, h, bg, fg, fs) {
  slide.addShape(pptx.ShapeType.roundRect, { x,y,w,h, fill:{color:bg}, line:{color:bg}, rectRadius:0.08 });
  slide.addText(text, { x,y,w,h, fontSize:fs||13, bold:true, color:fg||C.white, fontFace:'Calibri', align:'center', valign:'middle', wrap:true });
}

function dnarrow(s,x,y) {
  s.addShape(pptx.ShapeType.downArrow,  {x,y, w:0.4, h:0.32, fill:{color:C.muted}, line:{color:C.muted}});
}

// ---- SLIDE 1: Title ----------------------------------------------------------
{
  const s = pptx.addSlide();
  s.background = { color: C.asosBlack };
  // ASOS black & white style - no colour decoration, pure minimal
  s.addText('Carrier Booking Tool\nVBKREQ Generator', {
    x:0.72, y:1.08, w:9.18, h:2.5,
    fontSize:42, bold:true, color:C.white, fontFace:'Calibri Light',
    align:'left', valign:'top', wrap:true
  });
  s.addText('Supporting Test & React Suppliers to Raise Carrier Bookings - Without AIM', {
    x:0.72, y:3.75, w:9.18, h:0.6,
    fontSize:18, color:C.white, fontFace:'Calibri Light',
    align:'left', valign:'middle'
  });
  s.addText('ASOS - L&D Inbound Team  |  July 2026', {
    x:0.72, y:6.9, w:9.18, h:0.45,
    fontSize:11, color:C.asosGrey, fontFace:'Calibri', align:'left', valign:'middle'
  });
}

// ---- SLIDE 2: Why This Tool --------------------------------------------------
{
  const s = contentSlide('Why This Tool Exists');

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:1.15,w:5.9,h:5.6, fill:{color:'FFF1F2'}, line:{color:C.red,pt:2}, rectRadius:0.1});
  s.addText('[X]  The Problem', {x:0.45,y:1.25,w:5.7,h:0.5, fontSize:16, bold:true, color:C.red, fontFace:'Calibri'});
  s.addText(
    'AIM is the standard ASOS app for raising carrier booking requests (VBKREQs).\n\n' +
    'Suppliers find it difficult to use AIM for raising bookings.\n\n' +
    'Without a booking in the carrier system, teams have to manually track shipments through Excel - error-prone and time-consuming.',
    {x:0.5,y:1.8,w:5.65,h:4.7, fontSize:14, color:C.dark, fontFace:'Calibri', valign:'top', wrap:true}
  );

  s.addShape(pptx.ShapeType.roundRect, {x:6.6,y:1.15,w:6.38,h:5.6, fill:{color:C.greenBg}, line:{color:C.mint,pt:2}, rectRadius:0.1});
  s.addText('[OK]  The Solution', {x:6.7,y:1.25,w:6.1,h:0.5, fontSize:16, bold:true, color:C.mint, fontFace:'Calibri'});
  s.addText([
    {text:'Simple Excel template - ', options:{bold:true, fontSize:14, color:C.dark}},
    {text:'supplier fills it and sends by email. No AIM needed.\n\n', options:{fontSize:14, color:C.dark}},
    {text:'Automated pipeline - ', options:{bold:true, fontSize:14, color:C.dark}},
    {text:'tool fetches ASN & PO data from Databricks automatically.\n\n', options:{fontSize:14, color:C.dark}},
    {text:'VBKREQ XML generated ', options:{bold:true, fontSize:14, color:C.dark}},
    {text:'and sent directly to the carrier via SFTP.\n\n', options:{fontSize:14, color:C.dark}},
    {text:'Carrier tracks milestones ', options:{bold:true, fontSize:14, color:C.dark}},
    {text:'systematically end-to-end.', options:{fontSize:14, color:C.dark}},
  ], {x:6.7,y:1.8,w:6.1,h:4.8, fontFace:'Calibri', valign:'top', wrap:true});
}

// ---- SLIDE 3: System Architecture -------------------------------------------
{
  const s = contentSlide('System Architecture');

  // Row 1: 3 input sources (y=1.15, h=1.2)
  box(s, 'Supplier Email\n(Excel attached)', 0.35, 1.15, 3.6, 1.2, C.blue,   C.white, 13);
  box(s, 'Azure Databricks\n(ASN + PO data)', 4.85, 1.15, 3.6, 1.2, C.accent, C.white, 13);
  box(s, 'SharePoint\n(Graph API)',            9.35, 1.15, 3.62, 1.2, C.mint,   C.white, 13);

  // Cron note bar below input boxes (y=2.42, h=0.38) — no overlap
  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:2.42,w:12.62,h:0.38, fill:{color:'E0F2FE'}, line:{color:C.accent}, rectRadius:0.05});
  s.addText('Cron job (09:00 / 13:00): uploads the supplier email attachment to SharePoint automatically', {
    x:0.5,y:2.42,w:12.3,h:0.38, fontSize:12, color:C.blue, fontFace:'Calibri', italic:true, align:'center', valign:'middle'
  });

  // Down arrows from each input box (y=2.82)
  dnarrow(s, 1.9,  2.82);
  dnarrow(s, 6.45, 2.82);
  dnarrow(s, 11.0, 2.82);

  // Engine (y=3.18)
  box(s, 'VBKREQ Generator Engine  (Node.js / Express)\nBible Builder  |  VBKREQ Builder  |  SFTP Uploader', 0.35,3.18,12.62,1.2, C.navy, C.white, 14);

  // Down arrows to outputs (y=4.4)
  dnarrow(s, 1.15, 4.4);
  dnarrow(s, 4.6,  4.4);
  dnarrow(s, 8.05, 4.4);
  dnarrow(s, 11.5, 4.4);

  // Output boxes (y=4.74)
  box(s, 'VBKREQ XML\n-> E2open SFTP\n-> Davis Turner', 0.35, 4.74, 2.9, 1.45, C.amber,  C.white, 12);
  box(s, 'Booking Report\nEmail to Team',               3.75, 4.74, 2.9, 1.45, C.accent, C.white, 12);
  box(s, 'Generation Log\n(3-day rolling)',             7.2,  4.74, 2.9, 1.45, C.mint,   C.white, 12);
  box(s, 'Re-Submit /\nCancel (UI card)',              10.65, 4.74, 2.68,1.45, C.blue,   C.white, 12);
}

// ---- SLIDE 4: Workflow -------------------------------------------------------
{
  const s = contentSlide('End-to-End Booking Workflow');

  const steps = [
    {n:'1', label:'Supplier emails\nExcel template\nto ASOS team',       color:C.blue  },
    {n:'2', label:'Cron job uploads\nto SharePoint\n(09:00 / 13:00)',    color:C.accent},
    {n:'3', label:'Tool fetches\nASN + PO data\nfrom Databricks',       color:C.mint  },
    {n:'4', label:'Bible Build -\nmerge & validate\nmaster dataset',     color:C.blue  },
    {n:'5', label:'VBKREQ XML\ngenerated\n(New / Re-sub / Cancel)',      color:C.accent},
    {n:'6', label:'XML uploaded\nto E2open SFTP\n-> Carrier consumes',  color:C.amber },
  ];

  const bw=1.9, bh=2.0, by=1.3, gap=0.18;
  steps.forEach((st,i) => {
    const bx = 0.35 + i*(bw+gap);
    s.addShape(pptx.ShapeType.ellipse, {x:bx+bw/2-0.28, y:by-0.38, w:0.56,h:0.56, fill:{color:st.color}, line:{color:st.color}});
    s.addText(st.n, {x:bx+bw/2-0.28, y:by-0.38, w:0.56,h:0.56, fontSize:16, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:by,w:bw,h:bh, fill:{color:st.color}, line:{color:st.color}, rectRadius:0.1});
    s.addText(st.label, {x:bx,y:by,w:bw,h:bh, fontSize:13, color:C.white, fontFace:'Calibri', align:'center', valign:'middle', wrap:true});
    if (i < steps.length-1)
      s.addShape(pptx.ShapeType.rightArrow, {x:bx+bw, y:by+bh/2-0.14, w:gap, h:0.28, fill:{color:C.muted}, line:{color:C.muted}});
  });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:3.55,w:12.62,h:0.62, fill:{color:'FFFBEB'}, line:{color:C.amber}, rectRadius:0.07});
  s.addText('(!) Smart-Skip: Cancelled ASNs, cancelled POs, and already-booked shipments are automatically excluded - no VBKREQ raised.', {
    x:0.5,y:3.55,w:12.3,h:0.62, fontSize:13, color:C.amber, fontFace:'Calibri', bold:true, align:'left', valign:'middle', wrap:true
  });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:4.28,w:12.62,h:0.58, fill:{color:C.lightBg}, line:{color:C.accent}, rectRadius:0.07});
  s.addText('Corrections: Re-Submit (Cd 15) or Cancel (Cd 01) available from the UI - no template re-upload needed.', {
    x:0.5,y:4.28,w:12.3,h:0.58, fontSize:13, color:C.blue, fontFace:'Calibri', align:'left', valign:'middle', wrap:true
  });

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:4.97,w:12.62,h:0.58, fill:{color:C.greenBg}, line:{color:C.mint}, rectRadius:0.07});
  s.addText('After every run: a booking report email is sent to the team with VB Ref, PO numbers, carton counts, weight, and SFTP status.', {
    x:0.5,y:4.97,w:12.3,h:0.58, fontSize:13, color:C.mint, fontFace:'Calibri', align:'left', valign:'middle', wrap:true
  });
}

// ---- SLIDE 5: Stakeholders ---------------------------------------------------
{
  const s = contentSlide('Stakeholders & Roles');

  const cols = [
    { label:'Test & React\nSupplier',       color:C.blue,   icon:'Supplier',
      desc:'Sends completed Excel booking template by email.\nNo AIM access or XML knowledge required.' },
    { label:'ASOS L&D\nInbound Team',  color:C.accent, icon:'L&D Inbound',  
      desc:'Owns and operates the tool. Monitors the pipeline, handles exceptions, maintains the system.' },
    { label:'Davis Turner\n(Carrier)',       color:C.amber,  icon:'Carrier',
      desc:'Receives VBKREQ XML via E2open SFTP. Plans collection and tracks shipment milestones.' },
    { label:'E2open Platform',              color:C.muted,  icon:'E2open',
      desc:'EDI middleware routing VBKREQ messages from ASOS to the carrier and returning status events.' },
  ];

  const bw=2.9, bh=2.5, gap=0.37;
  cols.forEach((c,i) => {
    const bx = 0.35 + i*(bw+gap);
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:1.25,w:bw,h:bh, fill:{color:c.color}, line:{color:c.color}, rectRadius:0.1});
    s.addText(c.icon,   {x:bx,y:1.3, w:bw,h:0.6, fontSize:15, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
    s.addText(c.label,  {x:bx,y:1.9, w:bw,h:0.65, fontSize:13, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle', wrap:true});
    s.addShape(pptx.ShapeType.rect, {x:bx+0.04,y:2.58,w:bw-0.08,h:1.1, fill:{color:'F8FAFC'}, line:{color:'E2E8F0'}});
    s.addText(c.desc, {x:bx+0.08,y:2.58,w:bw-0.16,h:1.1, fontSize:12, color:C.dark, fontFace:'Calibri', align:'center', valign:'middle', wrap:true});
  });

  s.addShape(pptx.ShapeType.rightArrow, {x:0.35,y:3.9,w:12.62,h:0.38, fill:{color:C.accent}, line:{color:C.accent}});
  s.addText('Booking flow direction  ->', {x:0.35,y:4.35,w:12.62,h:0.3, fontSize:11, color:C.muted, fontFace:'Calibri', align:'center'});

  s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:4.75,w:12.62,h:0.55, fill:{color:C.lightBg}, line:{color:C.accent}, rectRadius:0.06});
  s.addText('ASOS Procurement team raises POs and the PO data is pulled automatically from Databricks - no manual input required from Procurement.', {
    x:0.5,y:4.75,w:12.3,h:0.55, fontSize:12.5, color:C.blue, fontFace:'Calibri', align:'left', valign:'middle', wrap:true
  });
}

// ---- SLIDE 6: Field Guide ----------------------------------------------------
{
  const s = contentSlide('Supplier Excel Template - Field Guide');

  const sections = [
    { y:1.12, bg:C.pink,   label:'MANDATORY (pink) - must be filled by the supplier for every PO row',
      body:'PO_Number  |  Cargo_Ready_Planned_Collection_Date (DD/MM/YYYY)  |  Carrier_Booking_Request_Date (DD/MM/YYYY)\n' +
           'Traffic_Mode (CFS or CY)  |  Booking_Group  |  No_of_Cartons  |  Unit_Weight_KG (kg per unit)  |  Carton_Type (dropdown)',
      bh:0.85 },
    { y:2.55, bg:'1B7F3E', label:'DEFAULTED (green) - pre-filled by the tool with sensible defaults; update only if different for your shipment',
      body:'Pack_Type = Bulk Flat     |     Collection_Type = Delivery     |     Hazardous = N/A',
      bh:0.44 },
    { y:3.55, bg:C.accent, label:'AUTO-FILLED (blue) - calculated automatically from Carton_Type via lookup; do NOT edit',
      body:'Carton_Length_cm  |  Carton_Width_cm  |  Carton_Height_cm  |  Carton_Weight_KG',
      bh:0.44 },
    { y:4.55, bg:C.muted,  label:'OPTIONAL (grey) - leave blank if not applicable',
      body:'Collection_Time (HH:MM)  |  Remarks',
      bh:0.44 },
    { y:5.55, bg:C.navy,   label:'Booking_Group - controls how POs are grouped into one VBKREQ',
      body:'"Single Booking" = one VBKREQ per PO  |  "Multiple POs-BK001..BK025" = grouped POs share one VBKREQ  |  "Multiple" = all POs in one VBKREQ',
      bh:0.44 },
  ];

  sections.forEach(sec => {
    s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:sec.y, w:12.62,h:0.42, fill:{color:sec.bg}, line:{color:sec.bg}, rectRadius:0.07});
    s.addText(sec.label, {x:0.45,y:sec.y, w:12.42,h:0.42, fontSize:14, bold:true, color:C.white, fontFace:'Calibri', valign:'middle'});
    s.addText(sec.body,  {x:0.5, y:sec.y+0.44, w:12.4, h:sec.bh, fontSize:13, color:C.dark, fontFace:'Calibri', valign:'top', wrap:true});
  });
}

// ---- SLIDE 7: VBKREQ Field Mapping -------------------------------------------
{
  const s = contentSlide('VBKREQ - Field Mapping & Sources');

  // Legend
  const legItems = [
    {label:'Supplier Template', color:C.pink},
    {label:'Template (defaulted)',color:'1B7F3E'},
    {label:'Databricks',        color:C.accent},
    {label:'Auto-generated',    color:C.mint},
  ];
  legItems.forEach((l,i) => {
    const lx = 0.35 + i*3.15;
    s.addShape(pptx.ShapeType.roundRect, {x:lx,y:1.08,w:0.18,h:0.28, fill:{color:l.color}, line:{color:l.color}, rectRadius:0.03});
    s.addText(l.label, {x:lx+0.22,y:1.08,w:2.8,h:0.28, fontSize:11, color:C.dark, fontFace:'Calibri', valign:'middle'});
  });

  // Table data: [VBKREQ Field, XML Element, Actual Value, source color]
  const tRows = [
    // -- Header / Reference fields --
    ['Transport Mode',              '<Mode>',                              '10 = Sea  |  30 = Road  |  40 = Air  |  50 = Rail  |  60 = Air-Supplier',  C.accent],
    ['Traffic Mode',                '<Reference RefTypeCd="QY">',          'CFS  |  CY',                                                               C.pink  ],
    ['Country of Origin',           '<Reference RefTypeCd="4B">',          'CN  |  IN  |  BD  |  TR  (bam033j PODtl[0].OriginCountryID)',              C.accent],
    ['Hazardous',                   '<Reference RefTypeCd="BH">',          'N  |  Y  (from Hazardous column; default = N/A)',                          '1B7F3E'],
    ['Customs Class',               '<Reference RefTypeCd="CC">',          'Green  (hardcoded - always Green)',                                         C.mint  ],
    ['Collection Type',             '<Reference RefTypeCd="CD">',          'Delivery  |  Collection  (default = Delivery)',                            '1B7F3E'],
    ['Collection Time (optional)',  '<Reference RefTypeCd="CT">',          '09:00  (required when Collection_Type = Collection)',                       C.pink  ],
    ['Booking Comments (optional)', '<Remark Qualifier="BRC">',            'Free text from Remarks column',                                            C.pink  ],
    // -- Trade Partners --
    ['Supplier Name + ID',          '<TradePartner RoleCd="SU">',          'Example Supplier Ltd  /  SUP001  (bam033j SupplierName + SupplierID)',     C.accent],
    ['Factory Name + ID',           '<TradePartner RoleCd="FA">',          'Factory Co Ltd  /  FAC123  (bam033j FactoryDesc + Factory)',               C.accent],
    ['Final Destination FC',        '<TradePartner RoleCd="FD">',          'FC01 Barnsley  /  FC01  (hardcoded FC address lookup)',                    C.mint  ],
    ['Carrier ID',                  '<TradePartner RoleCd="CA">',          '3  (Davis Turner - hardcoded)',                                            C.mint  ],
    ['Loading Port',                '<TradePartner RoleCd="SL">',          'CNSHA  |  INBOM  |  BDCGP  (bam033j LadingPort LOCODE)',                  C.accent],
    // -- Status Dates --
    ['Cargo Ready Date',            '<Status DateTypeCd="018">',           '2026-07-25  (Cargo_Ready_Planned_Collection_Date)',                        C.pink  ],
    ['Booking Request Date',        '<Status DateTypeCd="081">',           '2026-07-13  (Carrier_Booking_Request_Date)',                               C.pink  ],
    ['Ship Date (Ex-Factory)',      '<Status DateTypeCd="238">',           '2026-08-05  (bam033j ExFactoryDate minus 1 day)',                          C.accent],
    ['Expected Delivery',           '<Status DateTypeCd="065">',           '2026-08-20  (bam033j ExpectedDeliveryDateFirstLocation minus 1 day)',      C.accent],
    ['Destination FC LOCODE',       '<Status Location LocTypeCd="E/D">',   'GBBSY=FC01  |  GBLIC=FC02  |  GBHEM=FC03  |  DEBER=FC04',                 C.mint  ],
    ['System Timestamps',           '<Status DateTypeCd 211/OSBT/OSBK/SBK>','211 & SBK = current generation time.  OSBT & OSBK = original v1.0 booking time (preserved on re-submit/cancel)', C.mint  ],
    // -- Document totals --
    ['Booking Ref + Version',       '<Document> <Reference ACE/V0>',       'VB-000349  /  1.0  (auto-incrementing; 2.0 on re-submit)',                 C.mint  ],
    // Header-level measures (Document) — sourced from supplier template
    ['Header BKQ',                  '<Measure Qualifier="BKQ" UOMCd="UN">', 'Template: Total booked units of a booking  |  Fallback: sum of ASN line BKQs (bam036e)', C.pink],
    ['Header N (Net Weight)',        '<Measure Qualifier="N" UOMCd="KG">',  'Template: Total items weight of booking  (total net weight filled directly by supplier)',  C.pink],
    ['Header G (Gross Weight)',      '<Measure Qualifier="G" UOMCd="KG">',  'N  +  (Carton_Weight_KG x No_of_Cartons)  |  carton weight from CARTON_TYPES lookup',     C.mint],
    ['Header VOL',                  '<Measure Qualifier="VOL" UOMCd="M3">', '(L x W x H / 1,000,000) x No_of_Cartons  |  dimensions from CARTON_TYPES by Carton_Type',C.mint],
    ['Header QUR (Cartons)',        '<Measure Qualifier="QUR" UOMCd="CT">', 'Template: Total no. of Cartons of booking',                                C.pink  ],
    // Line-level measures (LineItem) — computed per SKU
    ['Line BKQ (per SKU)',          '<Measure Qualifier="BKQ" UOMCd="UN">', 'bam036e ASNInItem.unit_qty (primary)  |  bam033j PhysicalQtyOrdered (fallback)',          C.accent],
    ['Line N (per SKU)',            '<Measure Qualifier="N" UOMCd="KG">',  '(Template N / Header BKQ) x line BKQ  =  per-unit weight x line booking qty',             C.mint  ],
    ['Line G (per SKU)',            '<Measure Qualifier="G" UOMCd="KG">',  'Line N  +  Carton_Weight_KG x 1  (1 carton per SKU line)',                                C.mint  ],
    ['Line VOL (per SKU)',          '<Measure Qualifier="VOL" UOMCd="M3">', '(L x W x H / 1,000,000) x 1  (1 carton per SKU line)',                                   C.mint  ],
    ['Line QUR (per SKU)',          '<Measure Qualifier="QUR" UOMCd="CT">', '1.0000  (always 1 carton per SKU line)',                                                   C.mint  ],
  ];

  const colW = [2.85, 3.55, 5.7];
  const startX = 0.35, startY = 1.38, rowH = 0.232;

  // Header row
  const hdrLabels = ['VBKREQ Field', 'XML Element', 'Actual Value / Source'];
  hdrLabels.forEach((h,ci) => {
    const cx = startX + colW.slice(0,ci).reduce((a,b)=>a+b,0);
    s.addShape(pptx.ShapeType.rect, {x:cx,y:startY,w:colW[ci],h:0.36, fill:{color:C.navy}, line:{color:'FFFFFF'}});
    s.addText(h, {x:cx+0.05,y:startY,w:colW[ci]-0.1,h:0.36, fontSize:12, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
  });

  tRows.forEach((r,i) => {
    const ry = startY + 0.36 + i*rowH;
    const bg = i%2===0 ? 'FFFFFF' : 'F1F5F9';
    const srcColor = r[3];
    colW.forEach((cw,ci) => {
      const cx = startX + colW.slice(0,ci).reduce((a,b)=>a+b,0);
      s.addShape(pptx.ShapeType.rect, {x:cx,y:ry,w:cw,h:rowH, fill:{color:bg}, line:{color:'E2E8F0'}});
    });
    // Col 0: field name (bold, dark)
    s.addText(r[0], {x:startX+0.05, y:ry, w:colW[0]-0.1, h:rowH, fontSize:10.5, bold:true, color:C.dark, fontFace:'Calibri', valign:'middle'});
    // Col 1: XML element (blue, courier)
    const cx1 = startX+colW[0];
    s.addText(r[1], {x:cx1+0.05, y:ry, w:colW[1]-0.1, h:rowH, fontSize:9, color:C.blue, fontFace:'Courier New', valign:'middle', wrap:true});
    // Col 2: value + colored left border to indicate source
    const cx2 = startX+colW[0]+colW[1];
    s.addShape(pptx.ShapeType.rect, {x:cx2,y:ry,w:0.06,h:rowH, fill:{color:srcColor}, line:{color:srcColor}});
    s.addText(r[2], {x:cx2+0.1, y:ry, w:colW[2]-0.15, h:rowH, fontSize:9.5, color:C.dark, fontFace:'Calibri', valign:'middle', wrap:true});
  });

  // Line-item note
  const noteY = startY + 0.36 + tRows.length * rowH + 0.04;
  s.addShape(pptx.ShapeType.roundRect, {x:startX,y:noteY,w:12.62,h:0.35, fill:{color:C.lightBg}, line:{color:C.accent}, rectRadius:0.05});
  s.addText('Line-item level also includes per-SKU identifiers: ASN Ref (SI) | SKU (SK) | Colour (CL) | Size (IZ) | Pack Type (PAC) | Carton Type (98) | Product Style (PT) | Description (DSC) | FC per line (FS)', {
    x:startX+0.1, y:noteY, w:12.42, h:0.35, fontSize:10, color:C.blue, fontFace:'Calibri', valign:'middle', wrap:true
  });
}

// ---- SLIDE 8: Purpose Codes --------------------------------------------------
{
  const s = contentSlide('VBKREQ Purpose Codes - New, Re-Submit & Cancel');

  const pcs = [
    {code:'13', title:'New Submission', color:C.mint,
     when:'First-time booking for a PO or group of POs',
     how: 'Full pipeline: email -> SharePoint -> Databricks fetch -> VBKREQ generated -> SFTP upload', ver:'Version 1.0'},
    {code:'15', title:'Re-Submission',  color:C.amber,
     when:'A previous booking needs correction (e.g. date change, quantity update)',
     how: 'Look up VB Ref by PO in the Re-Submit / Cancel card -> resends stored master data with incremented version', ver:'Version increments'},
    {code:'01', title:'Cancellation',   color:C.red,
     when:'PO is cancelled or booking must be revoked',
     how: 'Look up VB Ref by PO in the Re-Submit / Cancel card -> sends cancellation VBKREQ; no template re-upload needed', ver:'Version unchanged'},
  ];

  pcs.forEach((pc,i) => {
    const by = 1.2 + i*1.85;
    s.addShape(pptx.ShapeType.roundRect, {x:0.35,y:by,w:1.1,h:1.6, fill:{color:pc.color}, line:{color:pc.color}, rectRadius:0.1});
    s.addText('Cd\n'+pc.code, {x:0.35,y:by,w:1.1,h:1.6, fontSize:20, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
    s.addText(pc.title, {x:1.6,y:by+0.05,w:11.3,h:0.42, fontSize:18, bold:true, color:pc.color, fontFace:'Calibri'});
    s.addText([
      {text:'When: ', options:{bold:true, color:C.dark, fontSize:13.5}},
      {text:pc.when,  options:{color:C.dark, fontSize:13.5}},
    ], {x:1.6,y:by+0.5,w:11.3,h:0.38, fontFace:'Calibri'});
    s.addText([
      {text:'How:  ', options:{bold:true, color:C.dark, fontSize:13}},
      {text:pc.how,   options:{color:C.dark, fontSize:13}},
    ], {x:1.6,y:by+0.88,w:11.3,h:0.5, fontFace:'Calibri', wrap:true});
    s.addShape(pptx.ShapeType.roundRect, {x:1.6,y:by+1.38,w:2.1,h:0.28, fill:{color:pc.color}, line:{color:pc.color}, rectRadius:0.05});
    s.addText(pc.ver, {x:1.6,y:by+1.38,w:2.1,h:0.28, fontSize:11, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'});
  });
}

// ---- SLIDE 9: Summary & Next Steps -------------------------------------------
{
  const s = contentSlide('Summary & Next Steps');

  const cards = [
    {label:'Email-driven',    color:C.blue,   desc:'Supplier emails Excel - no AIM or SharePoint upload needed.'},
    {label:'Fully Automated', color:C.accent, desc:'Cron -> Databricks -> VBKREQ XML -> SFTP - end to end.'},
    {label:'Built-in Safety', color:C.mint,   desc:'Smart-Skip blocks duplicates, cancelled POs & ASNs.'},
    {label:'Carrier Ready',   color:C.amber,  desc:'Davis Turner receives structured XML and tracks all milestones.'},
  ];

  cards.forEach((c,i) => {
    const col=i%2, row=Math.floor(i/2);
    const bx=0.35+col*3.2, by=1.2+row*1.8;
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:by,w:3.0,h:1.65, fill:{color:'FFFFFF'}, line:{color:c.color,pt:2}, rectRadius:0.1});
    s.addShape(pptx.ShapeType.roundRect, {x:bx,y:by,w:3.0,h:0.45, fill:{color:c.color}, line:{color:c.color}, rectRadius:0.1});
    s.addText(c.label, {x:bx+0.1,y:by,w:2.8,h:0.45, fontSize:13, bold:true, color:C.white, fontFace:'Calibri', align:'left', valign:'middle'});
    s.addText(c.desc,  {x:bx+0.1,y:by+0.5,w:2.8,h:1.05, fontSize:12.5, color:C.dark, fontFace:'Calibri', wrap:true, valign:'top'});
  });

  s.addShape(pptx.ShapeType.rect, {x:6.85,y:1.15,w:0.04,h:5.6, fill:{color:'E2E8F0'}, line:{color:'E2E8F0'}});
  s.addText('Next Steps', {x:7.1,y:1.15,w:5.88,h:0.45, fontSize:16, bold:true, color:C.navy, fontFace:'Calibri'});

  const nexts = [
    {who:'For Test & React Suppliers', color:C.blue, steps:[
      'Fill pink (mandatory) columns in the Excel template',
      'Email the completed template to Inboundservices@asos.com',
    ]},
    {who:'For ASOS L&D Inbound Team', color:C.accent, steps:[
      'Monitor pipeline UI for Smart-Skip warnings or SFTP failures',
      'Use Re-Submit / Cancel card for any corrections',
      'Review Booking History log daily',
    ]},
  ];

  let ny=1.65;
  nexts.forEach(n => {
    s.addShape(pptx.ShapeType.roundRect, {x:7.1,y:ny,w:5.88,h:0.36, fill:{color:n.color}, line:{color:n.color}, rectRadius:0.06});
    s.addText(n.who, {x:7.15,y:ny,w:5.78,h:0.36, fontSize:13, bold:true, color:C.white, fontFace:'Calibri', valign:'middle'});
    ny+=0.4;
    n.steps.forEach(step => {
      s.addText('- ' + step, {x:7.2,y:ny,w:5.73,h:0.38, fontSize:12.5, color:C.dark, fontFace:'Calibri', valign:'middle', wrap:true});
      ny+=0.42;
    });
    ny+=0.15;
  });

  s.addShape(pptx.ShapeType.roundRect, {x:7.1,y:ny,w:5.88,h:0.55, fill:{color:C.lightBg}, line:{color:C.accent}, rectRadius:0.06});
  s.addText('Questions? Contact: Inboundservices@asos.com', {
    x:7.2,y:ny,w:5.68,h:0.55, fontSize:12, color:C.blue, fontFace:'Calibri', valign:'middle', wrap:true
  });
}

// ---- SLIDE 10: Thank You ------------------------------------------------------
{
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect, {x:0,y:6.9,w:13.33,h:0.6, fill:{color:C.accent}, line:{color:C.accent}});
  s.addText('Thank You', {
    x:0.6,y:2.0,w:12.1,h:1.6, fontSize:54, bold:true, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'
  });
  s.addText('Carrier Booking Tool - VBKREQ Generator\nASOS L&D Inbound Team  |  July 2026', {
    x:0.6,y:3.8,w:12.1,h:1.0, fontSize:18, color:C.cyan, fontFace:'Calibri', align:'center', valign:'middle'
  });
  s.addText('For questions, contact: Inboundservices@asos.com', {
    x:0.6,y:5.1,w:12.1,h:0.6, fontSize:14, color:C.muted, fontFace:'Calibri', align:'center', valign:'middle'
  });
  s.addText('ASOS - L&D Inbound Team  |  July 2026', {
    x:0.6,y:6.9,w:12.1,h:0.6, fontSize:12, color:C.white, fontFace:'Calibri', align:'center', valign:'middle'
  });
}

const outFile = path.join(OUT, 'CarrierBookingTool_ProcessOverview.pptx');
pptx.writeFile({ fileName: outFile }).then(() => {
  console.log('Saved: ' + outFile);
}).catch(err => { console.error('Error:', err.message); process.exit(1); });