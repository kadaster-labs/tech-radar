/**
 * This script shows how table data can be represented as JSON.
 * The JSON data can then be given to other systems through Power Automate flows.
 * This script does not affect pre-existing data in the workbook.
 *
 * A table is represented in JSON as an array of objects, where each object represents one table row.
 * The objects keys are the column names, and the values are the cell values.
 */
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getTable('Content');

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(texts);
  }

  sortArray(returnObjects);

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));
  return returnObjects;
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let NUM_COLUMNS = 5; // name	ring	quadrant	isNew	description
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i];
      continue;
    }

    let object = {};
    for (let j = 0; j < NUM_COLUMNS; j++) {
      object[objectKeys[j]] = values[i][j];
    }

    objectArray.push(object as TableData);
  }

  return objectArray;
}

// Sort alphabetically by 'ring','name'. This has two reasons
// 1. techradar adds 'ring's in expanded view by alphabetically adding 'name's. Therefore, for each quadrant the order of rings might differ, which is undesirable
// 2. by keeping a consistent ordering, we'll have a consistent list in Git, which keeps diffs small. Even when the table in Excel is sorted in some arbitrary way
function sortArray(returnObjects: TableData[]) {
  const ringPriority = {
    "Adopt": 1,
    "Explore": 2,
    "Assess": 3,
    "Monitor": 4
  };

  returnObjects.sort((a: TableData, b: TableData) => {
    const nameA = a.name.toUpperCase();
    const nameB = b.name.toUpperCase();

    const ringA = a.ring;
    const ringB = b.ring;

    if (ringPriority[ringA] < ringPriority[ringB]) {
      return -1;
    } else if (ringPriority[ringA] > ringPriority[ringB]) {
      return 1;
    } else if (nameA > nameB) {
      return 1;
    } else if (nameA < nameB) {
      return -1;
    } else {
      return 0;
    }
  })
}

interface TableData {
  name: string;
  ring: string;
  quadrant: string;
  isNew: string;
  description: string;
}
