/**
 * This script shows how table data can be represented as JSON.
 * The JSON data can then be given to other systems through Power Automate flows.
 * This script does not affect pre-existing data in the workbook.
 *
 * A table is represented in JSON as an array of objects, where each object represents one table row.
 * The objects keys are the column names, and the values are the cell values.
 */
function main(workbook: ExcelScript.Workbook): TableData[] {
  const activeSheet = workbook.getActiveWorksheet();

  const tables = activeSheet.getTables();
  if (tables.length === 0) {
    console.log("No tables found in the active worksheet");
    return [];
  }

  const table = tables[0];
  console.log(`Using table: ${table.getName()} from sheet: ${activeSheet.getName()}`);

  const texts = table.getRange().getTexts();
  const values = table.getRange().getValues();

  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(texts, values);

    returnObjects.sort((a, b) => {
      const ringComparison = a.ring.localeCompare(b.ring);
      if (ringComparison === 0) {
        return a.name.localeCompare(b.name);
      }
      return ringComparison;
    });
  }

  console.log(JSON.stringify(returnObjects));
  return returnObjects;
}

function excelSerialToDate(serial: number): string {
  // Excel's epoch starts on 1900-01-01, but it incorrectly treats 1900 as a leap year,
  // so serial 1 = 1900-01-01 and we subtract 1 to correct for that bug.
  const excelEpoch = new Date(1899, 11, 30); // 1899-12-30
  const date = new Date(excelEpoch.getTime() + serial * 86400000);
  return `${date.getDate()}-${date.getMonth() + 1}-${date.getFullYear()}`;
}

function parseTextDate(text: string): string {
  // Handle formats like "04-03-26", "2026-03-02", "2-3-2026", etc.
  const parts = text.split(/[-\/]/);
  if (parts.length !== 3) return text; // Can't parse, return as-is

  let day: number, month: number, year: number;

  if (parts[0].length === 4) {
    // YYYY-MM-DD
    [year, month, day] = parts.map(Number);
  } else if (parts[2].length === 4) {
    // D-M-YYYY or M-D-YYYY — assuming D-M-YYYY given Dutch locale
    [day, month, year] = parts.map(Number);
  } else {
    // Two-digit year, e.g. "04-03-26" → assume D-M-YY
    [day, month, year] = parts.map(Number);
    year += year < 100 ? 2000 : 0;
  }

  if (!day || !month || !year) return text;
  return `${day}-${month}-${year}`;
}

function returnObjectFromValues(texts: string[][], values: (string | number | boolean)[][]): TableData[] {
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];

  if (texts.length > 0) {
    objectKeys = texts[0];
  }

  const dateColumnKey = "Laatst bijgewerkt";
  const dateColumnIndex = objectKeys.indexOf(dateColumnKey);

  for (let i = 1; i < texts.length; i++) {
    let object: { [key: string]: string } = {};

    for (let j = 0; j < objectKeys.length; j++) {
      if (j === dateColumnIndex) {
        const raw = values[i][j];
        if (typeof raw === "number") {
          // Real date cell — convert from Excel serial
          object[objectKeys[j]] = excelSerialToDate(raw);
        } else {
          // Text cell — normalize to D-M-YYYY
          object[objectKeys[j]] = parseTextDate(texts[i][j]);
        }
      } else {
        object[objectKeys[j]] = texts[i][j];
      }
    }

    objectArray.push(object as unknown as TableData);
  }

  return objectArray;
}

interface TableData {
  name: string;
  ring: string;
  quadrant: string;
  isNew?: string;
  status?: string;
  description: string;
  "Laatst bijgewerkt": string;
}
