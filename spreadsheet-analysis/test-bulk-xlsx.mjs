import fs from "node:fs/promises";

const root = "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/spreadsheet-analysis";
const data = JSON.parse(await fs.readFile(`${root}/flat-head-bulk-after.json`, "utf8"));
const source = data.results[0];
const record = {
  ...source,
  table: {
    ...source.table,
    headers: source.table.headers.slice(1),
    rows: source.table.rows.map((row) => row.slice(1)),
  },
};
const response = await fetch("http://localhost:8012/api/extract/bulk/xlsx", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ records: [record] }),
});
if (!response.ok) throw new Error(await response.text());
await fs.writeFile(`${root}/Flat_Head_Bulk_JSON_test.xlsx`, new Uint8Array(await response.arrayBuffer()));
console.log(`headers=${record.table.headers.length} rows=${record.table.rows.length}`);
