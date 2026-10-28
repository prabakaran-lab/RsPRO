import fs from "node:fs/promises";

const root = "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/spreadsheet-analysis";
const data = JSON.parse(await fs.readFile(`${root}/flat-head-bulk-after.json`, "utf8"));
const source = data.results[0].table;
const keep = source.headers
  .map((header, index) => ({ header, index }))
  .filter(({ header }) => !/article number|price/i.test(header));
const json = {
  headers: keep.map(({ header }) => header),
  rows: source.rows.map(row => keep.map(({ index }) => row[index])),
};
const response = await fetch("http://localhost:8012/api/extract/table/xlsx", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ json, filename: "Flat-Head-Round-Knurled-Steel.pdf" }),
});
if (!response.ok) throw new Error(await response.text());
await fs.writeFile(`${root}/Manual_Table_JSON_test.xlsx`, new Uint8Array(await response.arrayBuffer()));
console.log(`headers=${json.headers.length} rows=${json.rows.length}`);
