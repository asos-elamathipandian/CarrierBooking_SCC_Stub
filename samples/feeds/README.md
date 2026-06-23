# Feed Sources & Configuration

The backend supports two ASN data sources, controlled by the `ASN_SOURCE` variable in `.env`.

---

## Data Source: Databricks (default — `ASN_SOURCE=databricks`)

The primary source. Queries Azure Databricks SQL Warehouse using Azure AD credentials
(no PAT required — uses your existing `az login` session via `AzureCliCredential`).

### Databricks tables used

| Catalog / Schema | Table | Purpose |
|---|---|---|
| `supplychain.conformed` | `aim_shipment_detail_v1` | ASN / shipment data (carrier, mode, dates, SKU quantities, port of load) |
| `sourcingandbuying.conformed` | `bam033j_purchase_order_v1` | PO enrichment — supplier name/ID, factory, incoterms, EAN, size, colour per SKU |

### Fields sourced from Databricks

| VBKREQ field | Databricks column | Table |
|---|---|---|
| Supplier name | `SupplierName` | `bam033j_purchase_order_v1` |
| Supplier ID / code | `SupplierID` | `bam033j_purchase_order_v1` |
| Loading port (UN/LOCODE) | `LadingPort` | `bam033j_purchase_order_v1` |
| Shipping terms (Incoterms) | `FreightTermsDescription` | `bam033j_purchase_order_v1` |
| Factory code | `Factory` | `bam033j_purchase_order_v1` |
| Factory name | `FactoryDesc` | `bam033j_purchase_order_v1` |
| EAN barcode | `PODtl[].EANItemID` | `bam033j_purchase_order_v1` |
| Size | `PODtl[].SizeName` | `bam033j_purchase_order_v1` |
| Colour | `PODtl[].ColourName` | `bam033j_purchase_order_v1` |
| Product description | `PODtl[].OptionDescription` | `bam033j_purchase_order_v1` |
| Origin country | `PODtl[].OriginCountryID` | `bam033j_purchase_order_v1` |
| Transport mode code | `mode` | `aim_shipment_detail_v1` |
| Carrier | `carrier` | `aim_shipment_detail_v1` |
| Ship date | `asnEstimatedShipmentDate` | `aim_shipment_detail_v1` |
| Expected delivery date | `asnDeliveryDateFinalDest` | `aim_shipment_detail_v1` |

### Required `.env` settings

```
DATABRICKS_HOST=adb-2908786112690092.12.azuredatabricks.net
DATABRICKS_HTTP_PATH=/sql/1.0/warehouses/9d9de70087d062a5
ASN_SOURCE=databricks
```

Ensure you are logged in: `az login --tenant 4af8322c-80ee-4819-a9ce-863d5afbea1c`

---

## Data Source: Azure Blob Storage (`ASN_SOURCE=blob`)

Legacy fallback. Reads carrier ASN XML files uploaded to Azure Blob Storage.

### Required `.env` settings

```
AZURE_PO_FEED_CONNECTION_STRING=<SAS connection string>
ASN_SOURCE=blob
```

---

## Data Source: Local files (no env vars set)

For testing without any cloud dependency. When `AZURE_STORAGE_CONNECTION_STRING` /
`AZURE_BLOB_CONTAINER_NAME` are **not set** and `ASN_SOURCE` is not `databricks`,
the backend reads PO and ASN XML feeds from this folder (`samples/feeds/`).

### Naming convention

| Feed type | Filename format         | Example                     |
|-----------|-------------------------|-----------------------------|
| PO feed   | `PO_{PONumber}.xml`     | `PO_500034227415.xml`       |
| ASN feed  | `ASN_{ASNRef}.xml`      | `ASN_ASN-123456.xml`        |

The `{PONumber}` / `{ASNRef}` must exactly match the values in column
`PO_Number` / `ASN_Ref` of the supplier Excel you upload in Step 1.

### How to use

1. Drop your sample PO XML file here as e.g. `PO_500034227415.xml`
2. Drop your sample ASN XML file here as e.g. `ASN_ASN-123456.xml`
3. Start the server (`npm start`) — no Azure env vars needed
4. Upload supplier Excel → Fetch Feeds will read from this folder

---

## VBKREQ Field Mapping Reference

> **Default source is Databricks** (`ASN_SOURCE=databricks`). Supplier template columns are used as
> fallback only when the Databricks value is blank.

### Status Date Codes

| DateTypeCd | Default source (Databricks) | Fallback (supplier template) | Description |
|------------|----------------------------|------------------------------|-------------|
| `018` | — | `Cargo_Ready_Planned_Collection_Date` | Cargo ready / planned collection date |
| `081` | — | `Carrier_Booking_Request_Date` | Booking request date |
| `211` | System timestamp | System timestamp | Booking submission timestamp |
| `OSBT` | System timestamp | System timestamp | OSBT timestamp |
| `238` | `aim_shipment_detail_v1.asnEstimatedShipmentDate` | `Ship_Date` | Ship date |
| `065` | `aim_shipment_detail_v1.asnDeliveryDateFinalDest` | `Expected_Delivery_Date` | Expected delivery date |
| `OSBK` | System timestamp | System timestamp | OSBK timestamp |
| `SBK` | System timestamp | System timestamp | SBK timestamp |

### TradePartner Role Codes

| RoleCd | Level | Default source (Databricks) | Fallback (supplier template) | Description |
|--------|-------|-----------------------------|------------------------------|-------------|
| `SU` | Message | `bam033j_purchase_order_v1.SupplierName` / `SupplierID` | `Supplier_Name` / `Supplier_ID` | Supplier |
| `FA` | Message | `bam033j_purchase_order_v1.FactoryDesc` / `Factory` | `Factory_Name` / `Factory_ID` + address fields | Factory (address omitted if all fields blank) |
| `FD` | Message | `aim_shipment_detail_v1.finalDestination` | `fcId` — hardcoded FC01 or FC_MASTER | Final destination / FC |
| `CA` | Message | `aim_shipment_detail_v1.carrier` | `Carrier_ID` (default `3`) | Carrier |
| `SL` | Message | `bam033j_purchase_order_v1.LadingPort` (UN/LOCODE) | Loading port LOCODE from supplier template | Shipping location |
| `FS` | Line item | `aim_shipment_detail_v1.firstDestination` | `row.FC_ID` or booking-level `fcId` | Final store / FC per line |

### Transport Mode

| Field | Default source (Databricks) | Fallback | Notes |
|-------|-----------------------------|----------|-------|
| `Transport_Mode_Code` | `aim_shipment_detail_v1.mode` | Supplier template → hardcoded `30` | Numeric mode code (e.g. `10`=Air, `30`=Sea, `50`=Road) |
| `Shipping_Terms` | `bam033j_purchase_order_v1.FreightTermsDescription` | `Shipping_Terms` from template | Incoterms (e.g. `FOB`, `CIF`) |

### Line Item Attribute Codes

| AttributeTypeCd | Default source (Databricks) | Fallback (supplier template) | Description |
|-----------------|----------------------------|------------------------------|-------------|
| `SI` | `aim_shipment_detail_v1.asnId` | `ASN_Ref` | ASN reference |
| `SK` | `aim_shipment_detail_v1.orderLineItems[].sku` | `SKU` | SKU code |
| `CL` | `bam033j_purchase_order_v1.PODtl[].ColourName` | `Colour_Code` | Colour |
| `IZ` | `bam033j_purchase_order_v1.PODtl[].SizeName` | `Size_Code` | Size |

### Line Item Reference Codes

| RefTypeCd | Default source (Databricks) | Fallback (supplier template) | Description |
|-----------|----------------------------|------------------------------|-------------|
| `PAC` | `bam033j_purchase_order_v1.PODtl[].PackingMethod` | `Pack_Type` | Pack type (F=flat, H=hanging) |
| `PT` | `bam033j_purchase_order_v1.PODtl[].OptionItemID` | `Product_Style` | Product style |
| `DSC` | `bam033j_purchase_order_v1.PODtl[].OptionDescription` | `Description` | Product description |
| `HZ` | — | `Hazardous` | Hazardous flag |
| `98` | — | `Carton_Type` | Carton type |
| `LN` | — | `Carton_Length_cm` | Carton length (cm) |
| `WD` | — | `Carton_Width_cm` | Carton width (cm) |
| `HT` | — | `Carton_Height_cm` | Carton height (cm) |
