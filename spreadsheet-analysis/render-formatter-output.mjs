import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/tmp/formatter-flow-test.xlsx";
const outputDir = "C:/Users/Kalai/Videos/XSLX TO PDF/XSLX TO PDF/tmp/formatter-render";
await fs.mkdir(outputDir, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const renders = [
  ["Uniform Data", "A1:Q8", "uniform.png"],
  ["Source Backup", "A1:Q8", "source.png"],
  ["Mapping Guide", "A1:C18", "mapping.png"],
];

for (const [sheetName, range, name] of renders) {
  const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/${name}`, new Uint8Array(await preview.arrayBuffer()));
  console.log(`${sheetName}:${outputDir}/${name}`);
}
