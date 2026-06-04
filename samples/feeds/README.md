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
