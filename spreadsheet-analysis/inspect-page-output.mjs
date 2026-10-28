import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = process.argv[2] || "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/spreadsheet-analysis/Page_by_Page_Table_JSON_test.xlsx";
const sheetId = process.argv[3] || "PDF Ready JSON";
const range = process.argv[4] || "A1:D5";
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const result = await workbook.inspect({
  kind: "region",
  sheetId,
  range,
  maxChars: 2500,
  tableMaxRows: 5,
  tableMaxCols: 4,
  tableMaxCellChars: 60,
});
console.log(result.ndjson);
