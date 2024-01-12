import {
  getSheetBounds,
  getSheetCellValue,
  getSourceSheet,
} from "../functions";

function run() {
  const sheet = getSourceSheet(`${__dirname}\\sample1.xlsx`, "Sheet1");
  if (!sheet) return;

  console.log(__dirname);
  const bounds = getSheetBounds(sheet);
  console.log("iterating over a sheet with bounds", bounds);
  for (let y = bounds.from.y; y <= bounds.to.y; y++) {
    for (let x = bounds.from.x; x <= bounds.to.x; x++) {
      const val = getSheetCellValue(x, y, sheet);
      console.log(`Value at ${x}, ${y}`, val);
    }
  }
}

run();
