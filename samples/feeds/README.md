# Local Feed Files (for testing without Azure Blob Storage)

When `AZURE_STORAGE_CONNECTION_STRING` / `AZURE_BLOB_CONTAINER_NAME` env vars are **not set**,
the backend reads PO and ASN XML feeds from this folder.

## Naming convention

| Feed type | Filename format         | Example                     |
|-----------|-------------------------|-----------------------------|
| PO feed   | `PO_{PONumber}.xml`     | `PO_500034227415.xml`       |
| ASN feed  | `ASN_{ASNRef}.xml`      | `ASN_ASN-123456.xml`        |

The `{PONumber}` / `{ASNRef}` must exactly match the values in column
`PO_Number` / `ASN_Ref` of the supplier Excel you upload in Step 1.

## How to use

1. Drop your sample PO XML file here as e.g. `PO_500034227415.xml`
2. Drop your sample ASN XML file here as e.g. `ASN_ASN-123456.xml`
3. Start the server (`npm start`) — no Azure env vars needed
4. Upload supplier Excel → Fetch Feeds will read from this folder

---

## VBKREQ Field Mapping Reference

### Status Date Codes

| DateTypeCd | Source field | Description |
|------------|-------------|-------------|
| `018` | `Cargo_Ready_Planned_Collection_Date` | Cargo ready / planned collection date |
| `081` | `Carrier_Booking_Request_Date` | Booking request date |
| `211` | System timestamp | Booking submission timestamp |
| `OSBT` | System timestamp | OSBT timestamp |
| `238` | `Ship_Date` (carrier feed `<ShipDate>`) | Ship date |
| `065` | `Expected_Delivery_Date` (carrier feed `<ExpectedDeliveryDate>`) | Expected delivery date |
| `OSBK` | System timestamp | OSBK timestamp |
| `SBK` | System timestamp | SBK timestamp |

### TradePartner Role Codes

| RoleCd | Level | Source | Description |
|--------|-------|--------|-------------|
| `SU` | Message | `Supplier_Name` / `Supplier_ID` | Supplier |
| `FA` | Message | `Factory_Name` / `Factory_ID` + address fields | Factory (address omitted if all fields blank) |
| `FD` | Message | `fcId` — hardcoded lookup for FC01; dynamic FC_MASTER for others | Final destination / FC |
| `CA` | Message | `Carrier_ID` (default `3`) | Carrier |
| `SL` | Message | Loading port LOCODE | Shipping location |
| `FS` | Line item | `row.FC_ID` or booking-level `fcId` | Final store / FC per line |

### Line Item Attribute Codes

| AttributeTypeCd | Source | Description |
|-----------------|--------|-------------|
| `SI` | `ASN_Ref` | ASN reference |
| `SK` | `SKU` | SKU code |
| `CL` | `Colour_Code` | Colour code |
| `IZ` | `Size_Code` | Size code |

### Line Item Reference Codes

| RefTypeCd | Source | Description |
|-----------|--------|-------------|
| `PAC` | `Pack_Type` | Pack type |
| `PT` | `Product_Style` | Product style |
| `HZ` | `Hazardous` | Hazardous flag |
| `DSC` | `Description` | Description |
| `98` | `Carton_Type` | Carton type |
| `LN` | `Carton_Length_cm` | Carton length (cm) |
| `WD` | `Carton_Width_cm` | Carton width (cm) |
| `HT` | `Carton_Height_cm` | Carton height (cm) |
