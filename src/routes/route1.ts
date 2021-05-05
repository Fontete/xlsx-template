import { Router } from "express";
import { Request, Response } from "express";
import XlsxTemplate from "xlsx-template";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";
// import xlsxChart from "xlsx-chart";

// import {func1} from '../controllers/controller1'
const router1 = Router();

// router1.get("/", func1);
router1.post("/", (req: Request, res: Response) => {
  const workbook = xlsx.readFile("./template.xlsx");
  const sheet_name_list = workbook.SheetNames;
  // let sheet = {};
  // for (let i = 0; i < sheet_name_list.length; i++) {
  //   sheet = workbook.Sheets[sheet_name_list[i]];
  //   xlsx.utils.sheet_to_json(sheet, {
  //     header: "A",
  //     raw: true,
  //     blankrows: false,
  //     // defval: null,
  //   });
  // }
  const sheet1 = workbook.Sheets[sheet_name_list[0]];
  const sheet2 = workbook.Sheets[sheet_name_list[1]];
  const sheet3 = workbook.Sheets[sheet_name_list[2]];
  const sheet4 = workbook.Sheets[sheet_name_list[3]];
  const sheet5 = workbook.Sheets[sheet_name_list[4]];
  const sheet6 = workbook.Sheets[sheet_name_list[5]];
  const sheet7 = workbook.Sheets[sheet_name_list[6]];
  const sheet8 = workbook.Sheets[sheet_name_list[7]];
  const sheet9 = workbook.Sheets[sheet_name_list[8]];
  const sheet10 = workbook.Sheets[sheet_name_list[9]];

  const rawData = xlsx.utils.sheet_to_json(sheet3, {
    header: "A",
    raw: true,
    blankrows: false,
    // defval: null,
  });

  let newData: [] = [];
  for (let i = 0; i < rawData.length; i++) {
    if ((rawData[i].F && rawData[i].G) !== (null || undefined)) {
      const obj = { F: rawData[i].F, G: rawData[i].G };
      newData.push(obj);
    }
  }
  console.log(newData);

  const rawData2 = xlsx.utils.sheet_to_json(sheet4, {
    header: "A",
    raw: true,
    blankrows: false,
    defval: null,
  });

  const rawData3 = xlsx.utils.sheet_to_json(sheet5, {
    header: "A",
    raw: true,
    blankrows: false,
    defval: null,
  });
  //console.log("raw", rawData);

  fs.readFile(path.join(__dirname, "../../template.xlsx"), (err, data) => {
    const template: XlsxTemplate = new XlsxTemplate(data);
    const sheetNumber: number = 2;

    const values = {
      rawData,
      // newData,
      // rawData2,
      // rawData3,
      // A: [
      //   { A: "Đã hoàn thành" },
      //   { A: "Đang tiến hành" },
      //   { A: "Chưa thực hiện" },
      // ],
      // B: [{ B: 0.15 }, { B: 0.28 }, { B: 0.57 }],
    };
    console.log(values);

    template.substitute(sheetNumber, values);
    const result = template.generate({ type: "nodebuffer" });

    res.attachment("generateFile.xlsx");
    // xlsx.read(result, { type: "buffer" }).SheetNames.forEach((s) => {});
    // console.log("result", xlsx.read(result, { typC: "buffer" }).Sheets);
    res.send(result);
  });
});

export { router1 };
