import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = process.argv[2] || "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/spreadsheet-analysis/Page_by_Page_Table_JSON_test.xlsx";
const sheetName = process.argv[3] || "Extraction Review";
const range = process.argv[4] || "A1:F4";
const outputName = process.argv[5] || "pdf-ready-json.png";
const outputDir = "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/spreadsheet-analysis/bulk-render";
await fs.mkdir(outputDir, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
await fs.writeFile(`${outputDir}/${outputName}`, new Uint8Array(await preview.arrayBuffer()));
console.log(`${outputDir}/${outputName}`);
