# CarrierBookingStub — VBKREQ Generator

A web-based internal tool for ASOS to automate carrier booking requests (VBKREQs) to E2open/Davis Turner via SFTP. It eliminates manual XML authoring by pulling ASN and PO data from Azure Databricks, merging it with supplier-provided booking templates, and generating standards-compliant VBKREQ XML files ready for transmission.

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                        Browser UI                           │
│   (Vanilla JS + HTML — served by Express on port 3000)      │
└────────────────────────┬────────────────────────────────────┘
                         │ REST API calls
┌────────────────────────▼────────────────────────────────────┐
│                   Express Backend (Node.js)                  │
│                                                             │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────────┐  │
│  │ Supplier     │  │ Databricks   │  │  VBKREQ Builder   │  │
│  │ Template     │  │ ASN/PO       │  │  (XML generation) │  │
│  │ Parser       │  │ Reader       │  │                   │  │
│  └──────┬───────┘  └──────┬───────┘  └────────┬──────────┘  │
│         │                 │                   │             │
│         └────────── Bible Builder ────────────┘             │
│                    (master data merge)                      │
│                           │                                 │
│              ┌────────────▼───────────┐                     │
│              │    SFTP Uploader       │                     │
│              │  (E2open / DT SFTP)   │                     │
│              └────────────────────────┘                     │
│                                                             │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  SharePoint Scheduler (node-cron)                    │   │
│  │  Auto-fetches supplier template from SharePoint      │   │
│  │  at configured times (e.g. 09:00, 13:00 daily)      │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
         │                          │
┌────────▼────────┐      ┌──────────▼──────────┐
│ Azure Databricks│      │  SharePoint / Graph  │
│ (ASN + PO data) │      │  (Supplier templates)│
└─────────────────┘      └─────────────────────┘
```

---

## Key Features

| Feature | Description |
|---|---|
| **Supplier Template Upload** | Parse single-sheet Excel (PO Header) from suppliers; SKUs and quantities auto-booked from Databricks ASN feed |
| **SharePoint Auto-Sync** | Scheduled pull of the latest supplier Excel from a SharePoint folder (Graph API) |
| **ASN Enrichment** | Fetches shipment and PO detail from Azure Databricks (ADE) |
| **Bible Build** | Merges supplier template rows with ASN/PO data into a master dataset |
| **VBKREQ Generation** | Produces E2open-compliant XML with purpose codes: 13 (New), 15 (Re-Submission), 01 (Cancellation) |
| **Smart Skip** | No VBKREQ raised for cancelled ASNs (`_notification_type=C` in bam036e), cancelled POs (`Status=C`), or ASNs/POs that already have a carrier booking request (`bookingRequested` populated) |
| **SFTP Upload** | Transmits XML files directly to E2open/Davis Turner SFTP endpoint |
| **Re-Submit / Cancel** | Standalone card to look up previous VB Refs by PO and re-submit or cancel without re-uploading a template |
| **Booking History** | Rolling 3-day log of all generated VBKREQs with download links |
| **Local Mode** | SFTP_HOST can be left blank to save XML files locally for testing |

---

## Booking Workflow

```
1. Supplier uploads Excel template to SharePoint
          ↓
2. Auto-sync (09:00 / 13:00) or manual "Sync Now" fetches latest file
          ↓
3. Step 1 — Parse supplier template (PO refs extracted)
          ↓
4. Step 2 — Pipeline:
     a. Fetch ASN from Databricks (aim_shipment_detail_v1)
     b. Enrich with PO data (bam033j_purchase_order_v1)
     c. Build master dataset (Bible)
     d. Generate VBKREQ XML (purposeCd = 13)
     e. Upload to E2open SFTP
          ↓
5. Re-Submit (15) or Cancel (01) via standalone card if needed
```

---

## Purpose Codes

| Code | Meaning | Data Source |
|---|---|---|
| `13` | New Submission | Full pipeline (Databricks + supplier template) |
| `15` | Re-Submission | Stored master data from generation log |
| `01` | Cancellation | Stored master data from generation log |

---

## Technologies & Languages

### Languages
| Language | Usage |
|---|---|
| **JavaScript (Node.js)** | Backend server, all business logic |
| **JavaScript (Vanilla)** | Frontend UI (no framework) |
| **HTML / CSS** | Frontend markup and styling |
| **XML** | VBKREQ output format (E2open standard) |

### Runtime & Framework
| Software | Version | Purpose |
|---|---|---|
| **Node.js** | 18+ | Server runtime |
| **Express** | ^4.18 | HTTP server / REST API |

### Azure Services
| Service | Purpose |
|---|---|
| **Azure Databricks (ADE)** | Source of ASN and PO data (`aim_shipment_detail_v1`, `bam033j_purchase_order_v1`) |
| **Azure Blob Storage** | Legacy PO/ASN XML feed source (`bam033v`, `bam036` containers) |
| **Microsoft SharePoint** | Supplier template storage (auto-synced via Graph API) |
| **Microsoft Graph API** | SharePoint file access (`Sites.Read.All` app permission) |
| **Azure AD App Registration** | Service principal auth for Graph API (client credentials flow) |

### Key npm Packages
| Package | Purpose |
|---|---|
| `@azure/identity` | Azure AD client credentials authentication |
| `@azure/storage-blob` | Azure Blob Storage client (legacy feed) |
| `@databricks/sql` | Databricks SQL connector |
| `exceljs` | Parse supplier Excel templates (.xlsx) |
| `xmlbuilder2` | Build VBKREQ XML output |
| `xml2js` | Parse incoming XML feeds |
| `ssh2-sftp-client` | SFTP upload to E2open/Davis Turner |
| `node-cron` | Scheduled SharePoint sync (09:00, 13:00) |
| `multer` | Handle Excel file uploads |
| `dotenv` | Environment variable management |
| `express` | REST API server |

### External Systems
| System | Protocol | Purpose |
|---|---|---|
| **E2open / Davis Turner SFTP** | SFTP (port 22) | Destination for VBKREQ XML files |
| **SharePoint (asos1.sharepoint.com)** | HTTPS / Graph API | Supplier template source |
| **Azure Databricks** | HTTPS / Thrift SQL | ASN and PO enrichment data |

---

## Project Structure

```
CarrierBookingStub/
├── backend/
│   ├── server.js                  # Express app — all REST endpoints
│   ├── supplier-reader.js         # Parse supplier Excel templates
│   ├── bible-builder.js           # Merge supplier + ASN data into master dataset
│   ├── vbkreq-builder.js          # Generate VBKREQ XML
│   ├── databricks-asn-reader.js   # Fetch ASN/PO from Databricks
│   ├── sharepoint-client.js       # Graph API — list & download SharePoint files
│   ├── sharepoint-scheduler.js    # node-cron scheduled auto-sync
│   ├── sftp-uploader.js           # Upload XML to E2open SFTP
│   ├── blob-client.js             # Azure Blob Storage client (legacy)
│   ├── test-sharepoint.js         # Standalone SharePoint connection test
│   ├── test-sftp.js               # Standalone SFTP connection test
│   └── ctrl-counter.json          # Persistent VB Ref / version counters
├── frontend/
│   ├── index.html                 # Single-page UI
│   └── app.js                     # All frontend JavaScript
├── bible/
│   ├── generation-log.json        # History of all generated VBKREQs
│   ├── sp-sync-status.json        # SharePoint sync status
│   └── sharepoint-sync/           # Locally cached SharePoint downloads
├── output/                        # Generated VBKREQ XML files (local mode)
├── samples/
│   └── Supplier PO sheet-DDMMYYYY.xlsx  # Blank template to share with suppliers
├── config/
│   └── sftp.config.example.json   # SFTP config reference
├── .env                           # All credentials and configuration
└── package.json
```

---

## Setup & Running

### Prerequisites
- Node.js 18+
- Access to Azure Databricks warehouse
- E2open SFTP credentials
- Azure AD App Registration with `Sites.Read.All` (for SharePoint sync)

### Install
```bash
npm install
```

### Configure
Copy and fill in `.env` — key variables:
```
DATABRICKS_HOST / HTTP_PATH / TOKEN   ← Databricks SQL warehouse
SFTP_HOST / USERNAME / PRIVATE_KEY_PATH  ← E2open SFTP
SP_TENANT_ID / CLIENT_ID / CLIENT_SECRET / SITE_URL / FOLDER_PATH  ← SharePoint
SP_SCHEDULE=09:00,13:00               ← Auto-sync times
```

### Run
```bash
node backend/server.js
# or
npm start
```
App runs at **http://localhost:3000**

### Test connections
```bash
node backend/test-sharepoint.js   # Test SharePoint / Graph API
node backend/test-sftp.js         # Test E2open SFTP
```

---

## VBKREQ XML Notes

- Filename format: `DAVIESTN_E2ASOS_VBKREQ_1.0_{timestamp}{ctrlNumber}.xml`
- VB Refs format: `VB-{incrementingCounter}` (persisted in `ctrl-counter.json`)
- Version increments on re-submission (1.0 → 2.0 → 3.0…)
- Booking grouping: `Single Booking` (one VBKREQ per PO), `Multiple POs-BKxxx` (grouped), `Multiple` (all in one)
- Databricks dates are stored 1 day ahead — subtracted automatically before use
- **Cancelled / already-booked items are skipped automatically** — no VBKREQ is generated for:
  - ASNs where the latest `_notification_type` in `bam036e_asn_v1` is `C` (cancelled)
  - POs where `Status = C` in `bam033j_purchase_order_v1`
  - ASNs/POs where `bookingRequested` is already populated (booking already exists)
  - Skipped items are reported in the UI with reason details and do not block other POs from proceeding

---

## Booking Report Email

After each scheduled SharePoint sync (09:00 and 13:00) the tool automatically sends an HTML email summarising all carrier booking requests generated since the last report.

### Report columns

| Column | Description |
|--------|-------------|
| Supplier | Supplier name from Databricks PO feed |
| PO Number(s) | All POs in the booking |
| VB Ref | Carrier booking reference (VB-XXXXXX) |
| ASN Ref(s) | ASN IDs pulled from Databricks |
| Filename | Generated VBKREQ XML filename |
| Booking Group | Single Booking / Multiple POs-BKxxx / Multiple |
| Cargo Ready Date | Supplier-supplied cargo ready date |
| No. of Cartons | Total cartons across all POs in the booking |
| Total Weight (KG) | Total weight across all POs |
| SFTP Status | Uploaded / Pending |
| Generated At | Timestamp of VBKREQ generation |

### Trigger

The report fires once at the **end of each scheduled sync run**. It covers all VBKREQs generated since the previous report was sent (tracked in `bible/report-state.json`). **Manual UI bookings are not reported immediately** — they appear in the next scheduled report (09:00 or 13:00).

### Azure App Service — required setting

`node-cron` runs in-process. On Azure App Service the scheduler (and therefore the report) requires **Always On** to be enabled (Basic tier or above). Without it the Node.js process idles after ~20 minutes of inactivity and cron jobs stop firing.

> **Ephemeral filesystem note:** `bible/report-state.json` lives on the local App Service filesystem, which resets on restart or redeploy. After a restart the next report will include all historical log entries as a catch-up email rather than only the entries since the last run. To avoid this, promote `lastReportTime` to Azure Blob Storage or persist it as an app setting.

### Required App Registration permission

The sender mailbox (`REPORT_FROM`) needs the **`Mail.Send`** application permission in addition to the existing `Mail.ReadWrite`. In Azure AD: App Registration → API Permissions → Microsoft Graph → Application permissions → add `Mail.Send` → Grant admin consent.

### Configuration

```
REPORT_TO=InboundService@asos.com      # comma-separated recipients
REPORT_FROM=InboundService@asos.com    # sender mailbox (defaults to EMAIL_INGEST_MAILBOX)
```
