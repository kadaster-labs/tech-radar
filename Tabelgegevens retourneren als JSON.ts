/**
 * This script shows how table data can be represented as JSON.
 * The JSON data can then be given to other systems through Power Automate flows.
 * This script does not affect pre-existing data in the workbook.
 *
 * A table is represented in JSON as an array of objects, where each object represents one table row.
 * The objects keys are the column names, and the values are the cell values.
 */
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the active worksheet
  const activeSheet = workbook.getActiveWorksheet();

  // Get the first table in the active worksheet
  const tables = activeSheet.getTables();
  if (tables.length === 0) {
    console.log("No tables found in the active worksheet");
    return [];
  }

  const table = tables[0];
  console.log(`Using table: ${table.getName()} from sheet: ${activeSheet.getName()}`);

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(texts);

    // Sort by ring first, then by name
    returnObjects.sort((a, b) => {
      // First compare rings
      const ringComparison = a.ring.localeCompare(b.ring);

      // If rings are the same, compare names
      if (ringComparison === 0) {
        return a.name.localeCompare(b.name);
      }

      return ringComparison;
    });
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));
  return returnObjects;
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];

  // Get column headers from the first row
  if (values.length > 0) {
    objectKeys = values[0];
  }

  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      // Skip header row
      continue;
    }

    let object = {};
    for (let j = 0; j < objectKeys.length; j++) {
      object[objectKeys[j]] = values[i][j];
    }

    objectArray.push(object as TableData);
  }

  return objectArray;
}

interface TableData {
  name: string;
  ring: string;
  quadrant: string;
  isNew: string;
  description: string;
}
