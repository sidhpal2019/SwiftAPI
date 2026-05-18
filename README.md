# SwiftAPI

A fast, browser-based OpenAPI testing tool built with React. Originally migrated from the VBA-based ExcelAPI workflow.

## Implemented Functional Parity

- Load a Swagger/OpenAPI JSON definition file.
- Extract endpoint paths and available methods (GET/POST/DELETE/PUT/PATCH).
- Generate a test suite with sample parameter values, headers, and URL templates.
- Import test suite rows from Excel (.xlsx/.xls) and export current rows back to Excel.
- Edit param values, headers, and body before execution.
- Execute individual requests or run all rows.
- Capture response code, headers, response body, and response time in milliseconds.
- Optional OAuth2 client-credentials token retrieval and bearer injection.

## Run

```bash
cd SwiftAPI-app
npm install
npm run dev
```

Then open the local Vite URL shown in terminal.

## Notes

- Browser `fetch` is subject to CORS; if your APIs do not allow browser origins, use a proxy.
- OAuth2 credentials are not persisted and stay in browser state only.
- Existing VBA files remain untouched; this is a migration path, not a destructive replacement.

## Excel Columns Used

The import/export feature uses these headers in the workbook sheet (default sheet name: TestSuite):

- Protocol
- Server Path
- Path
- Method
- Params
- URL
- Param Values
- Headers
- Body
- Response Code
- Response Headers
- Response Body
- Response Time (ms)
