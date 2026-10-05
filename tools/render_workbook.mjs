import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const [input, sheet, range, output] = process.argv.slice(2);
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(input));
console.log((await wb.inspect({kind:'region', sheetId:sheet, range, maxChars:1500, tableMaxRows:4})).ndjson);
const blob = await wb.render({sheetName:sheet, range, scale:1.5, format:'png'});
await fs.writeFile(output, new Uint8Array(await blob.arrayBuffer()));
