import fs from "fs";
import xlsx, { WorkBook } from "xlsx";

type Obj = {
  [key: string]: any;
};

type DataFromWorkbook = {
  [key: string]: Obj[];
};

function getSourceData(path: string) {
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

function writeAsSheet(data: Obj[], filename: string) {
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
export function spreadsheetMap(
  inputPath: string,
  inputSheetName: string,
  outputName: string,
  mapFn: (item: Obj, i: number, array: Obj[]) => any
) {
  try {
    const sourceData = getSourceData(inputPath);
    const mapped = sourceData[inputSheetName].map(mapFn);
    writeAsSheet(mapped, outputName);
  } catch (error) {
    console.error(error);
  }
}

//this code produced test.xlsx
spreadsheetMap("sample1.xlsx", "Sheet1", "test", (item: Obj) => {
  const newObj = {
    id: item.id,
    name: item.name,
    startsWithLetter: item.name[0],
    is35: item.age === 35,
  };
  return newObj;
});
