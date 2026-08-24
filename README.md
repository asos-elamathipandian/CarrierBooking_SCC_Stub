# Carrier Booking Tool — Team Guide

> **What it does:** Automatically generates carrier booking requests (VBKREQs) to E2open / Davis Turner. Suppliers email a booking template → the tool pulls ASN & PO data from Databricks → produces compliant XML → uploads via SFTP. No manual XML authoring required.

---

## Contents

1. [Day-to-Day: How it works](#day-to-day-how-it-works)
2. [Using the UI](#using-the-ui)
3. [Adding a New Supplier](#adding-a-new-supplier)
4. [Troubleshooting](#troubleshooting)
5. [Azure Deployment — Updating the App](#azure-deployment--updating-the-app)
6. [Databricks Access — Raising a PR](#databricks-access--raising-a-pr)
7. [Environment Variables Reference](#environment-variables-reference)
8. [Local Development Setup](#local-development-setup)
9. [Architecture & Technical Reference](#architecture-overview)

---

## Day-to-Day: How it works

The tool runs **automatically** on Azure App Service. No one needs to manually trigger it.

```
Supplier emails booking template (.xlsx)
        ↓
Power Automate picks it up from InboundService@asos.com
        ↓
PA uploads file to Azure Blob Storage + triggers the App Service
        ↓
App Service parses the template, fetches ASN/PO data from Databricks
        ↓
Generates VBKREQ XML + uploads to E2open SFTP
        ↓
Sends summary email report to InboundService@asos.com at 09:00 and 13:00
```

### What your team needs to do day-to-day

| Scenario | Action |
|---|---|
| Supplier sends template | Nothing — PA handles it automatically |
| Check booking status | Open the tool UI → see pipeline status card |
| Booking failed | Check UI for error → see [Troubleshooting](#troubleshooting) |
| Re-submit a booking | UI → "Re-Submit / Cancel" card → enter PO → Re-Submit |
| Cancel a booking | UI → "Re-Submit / Cancel" card → enter PO → Cancel |
| Download a VBKREQ XML | UI → Booking History → Download |

**Tool URL:** `https://as-carrierbookingstub.azurewebsites.net`

---

## Using the UI

### Running the pipeline manually (if needed)

1. Open the tool URL above
2. **Upload supplier template** — drag & drop the `.xlsx` file from the supplier
3. Click **Run Pipeline** — the steps run in order:
   - Step 1: Parse supplier template
   - Step 2a: Fetch ASN from Databricks
   - Step 2b: Fetch PO enrichment data
   - Step 2c: Build master dataset (Bible)
   - Step 2d: Generate VBKREQ XML
   - Step 2e: Upload to SFTP
4. Green ticks = success. Amber ⚠ = some POs skipped (cancelled/already booked). Red = error.

### Skipped POs — what does ⚠ mean?

A PO is skipped (not sent to E2open) when:
- The ASN is **cancelled** (`D` status in Databricks)
- The PO is **cancelled** (`Status=C` in Databricks)
- A carrier booking **already exists** for that ASN/PO

This is expected behaviour. Expand the skipped section in the UI for details.

---

## Adding a New Supplier

### Standard supplier (same column headers as ASOS template)

No config needed — just upload their `.xlsx` file.

### Supplier with different column names (e.g. different header labels)

1. Copy the example mapping file:
   ```
   config/supplier-column-mapping.example.json → config/supplier-column-mapping.json
   ```
2. Edit `supplier-column-mapping.json` and add the supplier's column names mapped to the tool's canonical names:
   ```json
   {
     "PO_Number": ["PO Number", "Purchase Order"],
     "Carrier_Booking_Request_Date": ["Booking Request Date"]
   }
   ```
3. Save and restart the app. The parser picks it up automatically.

**Minimum required fields to map:**
`PO_Number`, `Cargo_Ready_Planned_Collection_Date`, `Carrier_Booking_Request_Date`, `Booking_Group`, `No_of_Cartons`, `Unit_Weight_KG`

### ASN-based supplier (e.g. Ideateks — template has ASN numbers, not PO numbers)

Use `ASN_Number` in the mapping instead of `PO_Number`. The tool automatically resolves PO numbers from ASN IDs via Databricks.

Ideateks defaults applied automatically:
- `Booking_Group` → `Single Booking`
- `Cargo_Ready_Planned_Collection_Date` → today
- `Carrier_Booking_Request_Date` → tomorrow
- `Traffic_Mode` → `CFS`

---

## Troubleshooting

### "Databricks 403 / PERMISSION_DENIED"

The Azure App Service SP doesn't have access to Databricks. See [Databricks Access — Raising a PR](#databricks-access--raising-a-pr).

### "SFTP upload failed"

- Check SFTP credentials in Azure App Service → Configuration → `SFTP_HOST`, `SFTP_USERNAME`, `SFTP_PRIVATE_KEY_PATH`
- Run connection test locally: `node backend/test-sftp.js`

### "SharePoint sync failed — 403"

- The `Sites.Selected` permission may have expired or not been applied per-site
- Check App Registration `carrier-booking-sharepoint-write` in Azure Portal → API Permissions
- Ensure admin consent has been granted AND the per-site write grant was applied via Graph API

### "Pipeline runs locally but fails on Azure"

Most likely cause: the SP (`carrier-booking-sharepoint-write`) is missing a permission in one of the downstream services (Databricks, SharePoint, or Graph). Check Azure App Service logs:
1. Azure Portal → App Service `as-carrierbookingstub` → Log stream
2. Or: App Service → Advanced Tools (Kudu) → Debug Console → `LogFiles/`

### "Emails not sending / report not arriving"

- Check `Mail.Send` application permission is granted on the App Registration
- Confirm `REPORT_TO` and `REPORT_FROM` env vars are set in App Service Configuration
- Check **Always On** is enabled (App Service → Configuration → General settings) — without it, the 09:00/13:00 cron jobs stop after 20 min of inactivity

### "All bookings re-sent after restart"

Known behaviour: `bible/report-state.json` lives on the App Service filesystem which resets on restart. The next report will be a catch-up of all historical entries. This is harmless — E2open deduplicates by VB Ref.

---

## Azure Deployment — Updating the App

The app runs on **Azure App Service `as-carrierbookingstub`** (Resource Group: `as-inbound-e2e-testsupporttool`).

### Deploying code changes

```bash
# From the repo root — zip deploy via Azure CLI
az webapp deployment source config-zip \
  --resource-group as-inbound-e2e-testsupporttool \
  --name as-carrierbookingstub \
  --src <path-to-zip>
```

Or use the Azure Portal → App Service → Deployment Center if connected to this repo.

### Adding / changing environment variables

Azure Portal → App Service `as-carrierbookingstub` → **Configuration** → Application settings → Add/Edit → Save → Restart.

> Never put secrets in code or commit `.env` to git. All secrets live in App Service Configuration only.

### Checking the app is running

Azure Portal → App Service → Overview → check Status = **Running**.  
Live logs: App Service → **Log stream**.

---

## Databricks Access — Raising a PR

If the SP needs access to new Databricks serve layer tables, raise a PR to:
[`asosteam/asos-data-ade-consumptionzone-infrastructure`](https://github.com/asosteam/asos-data-ade-consumptionzone-infrastructure)

### File to create
`terraform/phase2_databricks_config/modules/systematic_access/configs/prd/<system_name>.yaml`

### Template
```yaml
system_config:
  system_name: "e2e_carrier_booking"
  system_type: "business_data_consumer"
  description: "Data access for E2E Carrier Booking tool"

  owner:
    name: "e2e_carrier_booking"
    contact_email: "InboundService@asos.com"

  infra_config:
    cluster_size: "Small"
    auto_stop_minutes: 1
    enable_serverless: true
    max_num_clusters: 1
    enable_photon: false
    tags:
      environment: "Production"
      backstage_system: "<registered-backstage-system>"
      backstage_component: "<registered-backstage-component>"

  spns:
    - client_id: "322bfb99-27c5-444b-8e68-92612c73d5d9"
      display_name: "carrier-booking-sharepoint-write"
      entra_groups:
        - "AAD_ADE_SBU_Serve_System_DataReader"   # sourcingandbuying.serve
        - "AAD_ADE_SCH_Serve_System_DataReader"   # supplychain.serve
```

### Branch naming rule
Branch names must match: `feature/<6-7 digit ticket>-<10+ letters>`  
Example: `feature/2608211-AddCarrierBookingServeAccess`

### Reviewer
Tag **Japji Gandhi** (`asosjapji`) as reviewer.

---

## Capacity & Known Limits

| Constraint | Limit | Notes |
|---|---|---|
| Supplier Excel file size | 10 MB per file | Easily fits a typical booking template |
| Files per upload | 20 files | Hard limit in the upload handler |
| POs per Databricks query | ~500 recommended | All POs sent in a single SQL `IN (...)` clause — no batching |
| Session state | Azure App Service RAM | All pipeline data is in-memory; large datasets on Basic tier may cause issues |

The tool is designed for **day-to-day operational use** — typically 10–100 POs per supplier template. If you ever need 500+ POs in a single run, the Databricks query in `backend/databricks-asn-reader.js` would need batching logic added.

---

## Environment Variables Reference

All variables go in `.env` locally, or Azure App Service → Configuration on Azure.

| Variable | Required | Description |
|---|---|---|
| `DATABRICKS_HOST` | Yes | e.g. `adb-2908786112690092.12.azuredatabricks.net` |
| `DATABRICKS_HTTP_PATH` | Yes | SQL warehouse HTTP path |
| `DATABRICKS_TOKEN` | Local only | Personal access token (local dev only) |
| `SP_CLIENT_ID` | Yes | App Registration client ID (`322bfb99-...`) |
| `SP_CLIENT_SECRET` | Yes | App Registration client secret |
| `SP_TENANT_ID` | Yes | Azure AD tenant ID |
| `DATABRICKS_SKIP_SP` | Local only | Set `true` locally to use personal az login instead of SP |
| `SFTP_HOST` | Yes | E2open SFTP hostname |
| `SFTP_USERNAME` | Yes | SFTP username |
| `SFTP_PRIVATE_KEY_PATH` | Yes | Path to `.ppk` / `.pem` private key file |
| `SFTP_REMOTE_PATH` | Yes | Remote SFTP directory |
| `SP_SITE_URL` | Yes | SharePoint site URL |
| `SP_FOLDER_PATH` | Yes | SharePoint folder path for supplier templates |
| `SP_SCHEDULE` | No | Cron sync times, default `09:00,13:00` |
| `REPORT_TO` | Yes | Email recipients for booking report (comma-separated) |
| `REPORT_FROM` | No | Sender address, defaults to `EMAIL_INGEST_MAILBOX` |
| `EMAIL_INGEST_MAILBOX` | Yes | Mailbox to scan for supplier emails |
| `EMAIL_READONLY_MODE` | No | `true` = use watermark (default); `false` = mark as read |
| `AZURE_WEBHOOK_STORAGE_CONNECTION_STRING` | Yes | Blob storage connection string |
| `AZURE_WEBHOOK_CONTAINER_NAME` | No | Blob container name, default `supplier-uploads` |
| `NODE_ENV` | Azure only | Set to `production` on Azure (auto-set via web.config) |

---

## Local Development Setup

### Prerequisites
- Node.js 18+
- Azure CLI (`az login` with your ASOS account)
- Access to Azure Databricks warehouse `9d9de70087d062a5`

### First-time setup

```bash
git clone <repo-url>
cd CarrierBookingStub
npm install
```

Create `.env` from the reference above. For local dev, set:
```
DATABRICKS_SKIP_SP=true   # uses your personal az login, not the SP
SFTP_HOST=                # leave blank to save XML locally instead of uploading
```

### Run locally

```bash
node backend/server.js
# or
npm start
```

Open **http://localhost:3000**

### Test individual connections

```bash
node backend/test-sharepoint.js   # test SharePoint / Graph API
node backend/test-sftp.js         # test E2open SFTP
node backend/test-databricks.js   # test Databricks connection
```

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
| **Supplier Template Upload** | Parse single-sheet Excel (PO Header) emailed by suppliers to InboundService@asos.com; mandatory fields: PO_Number, Booking_Group, dates, Total booked units, cartons, weight; defaulted: Carton_Type (BDCM1), Pack_Type, Collection_Type, Hazardous, Traffic_Mode (CFS) |
| **SharePoint Auto-Sync** | Scheduled pull of the latest supplier Excel from a SharePoint folder (Graph API) |
| **ASN Enrichment** | Fetches shipment and PO detail from Azure Databricks (ADE) |
| **Bible Build** | Merges supplier template rows with ASN/PO data into a master dataset |
| **VBKREQ Generation** | Produces E2open-compliant XML with purpose codes: 13 (New), 15 (Re-Submission), 01 (Cancellation) |
| **Smart Skip** | No VBKREQ raised for cancelled ASNs (`_notification_type=D` in bam036e), cancelled POs (`Status=C`), or ASNs/POs that already have a carrier booking request (`bookingRequested` populated) |
| **SFTP Upload** | Transmits XML files directly to E2open/Davis Turner SFTP endpoint |
| **Re-Submit / Cancel** | Standalone card to look up previous VB Refs by PO and re-submit or cancel without re-uploading a template |
| **Booking History** | Rolling 3-day log of all generated VBKREQs with download links |
| **Local Mode** | SFTP_HOST can be left blank to save XML files locally for testing |

---

## Booking Workflow

```
1. Supplier emails Excel template to InboundService@asos.com
          ↓
2. Scheduled cron job (09:00 / 13:00) uploads email attachment to SharePoint
          ↓
3. Auto-sync fetches latest file from SharePoint
          ↓
4. Step 1 — Parse supplier template (PO refs, booking qty, carton data extracted)
          ↓
5. Step 2 — Pipeline:
     a. Fetch ASN from Databricks (aim_shipment_detail_v1)
     b. Enrich with PO data + per-SKU quantities (bam033j_purchase_order_v1)
     c. Fetch ASN cancellation status + booked unit_qty (bam036e_asn_v1)
     d. Build master dataset (Bible)
     e. Generate VBKREQ XML (purposeCd = 13)
        - Header BKQ = supplier "Total booked units" field (flags discrepancy vs ASN line sum)
        - Line-level N, G, VOL, QUR computed from carton type dimensions
     f. Upload to E2open SFTP
          ↓
6. Re-Submit (15) or Cancel (01) via standalone card if needed
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
| **Azure Databricks (ADE)** | Source of ASN and PO data (`aim_shipment_detail_v1`, `bam033j_purchase_order_v1`, `bam036e_asn_v1`) |
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
│   └── supplier-column-mapping.example.json # Optional column alias mapping for non-standard supplier templates
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

### Optional: Alternate Supplier Template Mapping (e.g. Ideateks)
If a supplier uses different Excel column names, you can map those headers to the tool's canonical fields.

1. Copy `config/supplier-column-mapping.example.json` to `config/supplier-column-mapping.json`.
2. Add aliases for each canonical field you receive from the supplier file.
3. Restart the server.

By default, the parser loads `config/supplier-column-mapping.json` automatically when present.
You can also point to a custom file path with:

```
SUPPLIER_COLUMN_MAPPING_FILE=./config/supplier-column-mapping.json
```

Supported JSON formats:

```json
{
  "PO_Number": ["PO Number", "PO No"],
  "Carrier_Booking_Request_Date": ["Booking Request Date"]
}
```

or

```json
{
  "PO Number": "PO_Number",
  "Booking Request Date": "Carrier_Booking_Request_Date"
}
```

Minimum mapped fields for header-only files are:
`PO_Number`, `Cargo_Ready_Planned_Collection_Date`, `Carrier_Booking_Request_Date`, `Booking_Group`, `No_of_Cartons`, `Unit_Weight_KG`.

For ASN-based files (for example Ideateks), use `ASN_Number` instead of `PO_Number`.
The backend resolves POs from ASN IDs using Databricks automatically.

Ideateks defaults applied when `ASN_Number` is used:
- `Booking_Group` forced to `Single Booking`
- `Cargo_Ready_Planned_Collection_Date` defaults to system date (today)
- `Carrier_Booking_Request_Date` defaults to next day
- `Traffic_Mode` defaults to `CFS` if not supplied

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

## Databricks Tables (Serve Layer)

The tool queries the ADE serve layer via a single star-schema query.

### Central fact table
**`sourcingandbuying.serve.fact_purchase_order_commitment_v1`**  
Grain: PO + ASN + SKU (snapshot). Latest row per PO+ASN+SKU selected via `ROW_NUMBER OVER PARTITION`.

| Field | Usage |
|---|---|
| `quantity` | Booking_Qty (BKQ) per SKU |
| `asn_status_code = 'D'` | Smart Skip — ASN cancelled |
| `is_booked_by_carrier = 'Yes'` | Smart Skip — booking already exists |

### Dimension tables joined

| Table | Key fields |
|---|---|
| `supplychain.serve.dim_advanced_shipment_notice_v1` | `asn_id`, `lading_port_code`, `mode_of_transport`, ship/delivery dates |
| `sourcingandbuying.serve.dim_supplier_v1` | `supplier_id`, `supplier` (name) |
| `sourcingandbuying.serve.dim_factory_v1` | `factory_id`, `factory_name`, `factory_country_code` |
| `sourcingandbuying.serve.dim_purchase_order_v1` | `po_number`, `inco_terms`, `origin_country_code` |

> Dates in the fact table are native `date` type — no −1 day correction needed. `1900-01-01` = null sentinel.

---

## VBKREQ XML Notes

- Filename format: `DAVIESTN_E2ASOS_VBKREQ_1.0_{timestamp}{ctrlNumber}.xml`
- VB Refs format: `VB-{incrementingCounter}` (persisted in `ctrl-counter.json`)
- Version increments on re-submission (1.0 → 2.0 → 3.0…)
- Booking grouping: `Single Booking` (one VBKREQ per PO), `Multiple POs-BKxxx` (grouped), `Multiple` (all in one)
- Databricks dates are stored 1 day ahead — subtracted automatically before use
- **Cancelled / already-booked items are skipped automatically** — no VBKREQ is generated for:
  - ASNs where the latest `_notification_type` in `bam036e_asn_v1` is `D` (deleted/cancelled)
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
