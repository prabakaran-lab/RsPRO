import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = process.argv[2] || "C:/Users/Kalai/Downloads/RS PRO POC-PDF Creation_ Grouping_363 SKUs_Batch 4_Part-2_05182026_Input (2).xlsx";
const input = await FileBlob.load(source);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table,region",
  maxChars: 16000,
  tableMaxRows: 8,
  tableMaxCols: 30,
  tableMaxCellChars: 120,
});

console.log(summary.ndjson);
