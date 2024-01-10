import fs from "fs";
import xlsx from "xlsx";

function getSourceData(path: string) {
  try {
    const file = fs.readFileSync(path);
    const workbook = xlsx.read(file, { type: "buffer" });
    const data: any = {};
    for (const sheetName of workbook.SheetNames) {
      data[`${sheetName}`] = xlsx.utils.sheet_to_json(
        workbook.Sheets[sheetName]
      );
    }

    return data;
  } catch (error) {
    console.error("ERROR PARSING SOURCE: ", error);
  }
}

console.log(getSourceData("./sample1.xlsx"));
