import { Router } from "express";
import { Request, Response } from "express";
import XlsxTemplate from "xlsx-template";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";
// import {func1} from '../controllers/controller1'
const router1 = Router();

// router1.get("/", func1);
router1.post("/", (req: Request, res: Response) => {
  const workbook = xlsx.readFile("./raw.xlsx");
  const sheet_name_list = workbook.SheetNames;
  // type inputData = Object[];
  const sheet1 = workbook.Sheets[sheet_name_list[0]];
  const rawData = xlsx.utils.sheet_to_json(sheet1, {
    header: "A",
    //blankrows: false,
  });
  console.log("raw", rawData);
  // const rawData = [
  //   {
  //     loivipham: "Bố trí phương tiện",
  //     trogia: [2000, 3000],
  //     ktg: [1000, 2000],
  //   },
  //   {
  //     loivipham: "lạng lách",
  //     trogia: [5000, 6000],
  //     ktg: [1000, 2000],
  //   },
  // ];

  fs.readFile(path.join(__dirname, "../../template.xlsx"), (err, data) => {
    const template: XlsxTemplate = new XlsxTemplate(data);
    const sheetNumber: number = 1;

    const values = {
      // extractDate: new Date(),
      // dates: [
      //   new Date("2013-06-01"),
      //   new Date("2013-06-02"),
      //   new Date("2013-06-03"),
      // ],
      rawData,
    };
    console.log(values);

    template.substitute(sheetNumber, values);
    const result = template.generate({ type: "nodebuffer" });

    res.attachment("generateFile.xlsx");
    // xlsx.read(result, { type: "buffer" }).SheetNames.forEach((s) => {});
    // console.log("result", xlsx.read(result, { type: "buffer" }).Sheets.Sheet1);
    res.send(result);
  });
});

export { router1 };
