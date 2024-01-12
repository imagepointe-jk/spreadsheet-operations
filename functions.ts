import fs from "fs";
import xlsx, { WorkBook, WorkSheet } from "xlsx";

type Obj = {
  [key: string]: any;
};

type DataFromWorkbook = {
  [key: string]: Obj[];
};

export function getSourceSheet(path: string, sheetName: string) {
  try {
    const file = fs.readFileSync(path);
    const workbook = xlsx.read(file, { type: "buffer" });
    return workbook.Sheets[sheetName];
  } catch (error) {
    console.error(error);
  }
}

export function getSourceJson(path: string) {
  try {
    const file = fs.readFileSync(path);
    const workbook = xlsx.read(file, { type: "buffer" });
    const data: DataFromWorkbook = {};
    for (const sheetName of workbook.SheetNames) {
      data[`${sheetName}`] = xlsx.utils.sheet_to_json(
        workbook.Sheets[sheetName]
      );
    }

    return data;
  } catch (error) {
    throw new Error("ERROR PARSING SOURCE");
  }
}

export function writeAsSheet(data: Obj[], filename: string) {
  const sheet = xlsx.utils.json_to_sheet(data);
  const workbook: WorkBook = {
    Sheets: {
      Sheet1: sheet,
    },
    SheetNames: ["Sheet1"],
  };
  xlsx.writeFile(workbook, `${filename}.xlsx`);
}

//like an array map, but instead of producing a new array, it produces a new spreadsheet using the given map function.
//iterates over each row, treating each one as an object. does NOT iterate cell-by-cell.
export function spreadsheetRowMap(
  inputPath: string,
  inputSheetName: string,
  outputName: string,
  mapFn: (row: Obj, i: number, array: Obj[]) => any
) {
  try {
    const sourceData = getSourceJson(inputPath);
    const mapped = sourceData[inputSheetName].map(mapFn);
    writeAsSheet(mapped, outputName);
  } catch (error) {
    console.error(error);
  }
}

function columnLetterLabelToNumber(label: string) {
  return label
    .split("")
    .reduce(
      (accum, letter, i, arr) =>
        accum + (letter.charCodeAt(0) - 64) * Math.pow(26, arr.length - 1 - i),
      0
    );
}

function numberToColumnLetterLabel(num: number) {
  let label = "";
  while (num > 0) {
    label = String.fromCharCode((num % 26) + 64) + label;
    num = Math.floor(num / 26);
  }
  return label;
}

export function getSheetCellValue(
  colNum: number,
  rowNum: number,
  sheet: WorkSheet
) {
  const colLabel = numberToColumnLetterLabel(colNum);
  const cellLabel = `${colLabel}${rowNum}`;
  return sheet[cellLabel].v;
}

export function getSheetBounds(sheet: WorkSheet) {
  const range = sheet["!ref"];
  if (!range) throw new Error("The sheet's range is not recognized.");

  const rangeSplit = range.split(":");
  const firstCell = rangeSplit[0];
  const lastCell = rangeSplit[1];
  const firstRowNumber = +firstCell.replace(/[^\d]/g, "");
  const lastRowNumber = +lastCell.replace(/[^\d]/g, "");
  const firstColumnLabel = firstCell.replace(/\d/g, "");
  const lastColumnLabel = lastCell.replace(/\d/g, "");
  const firstColumnNumber = columnLetterLabelToNumber(firstColumnLabel);
  const lastColumnNumber = columnLetterLabelToNumber(lastColumnLabel);

  return {
    from: {
      x: firstColumnNumber,
      y: firstRowNumber,
    },
    to: {
      x: lastColumnNumber,
      y: lastRowNumber,
    },
  };
}
