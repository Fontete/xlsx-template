import { Request, Response } from "express";
import XlsxTemplate from "xlsx-template";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";
import cloudinary, { ConfigOptions } from "cloudinary";

// config
// const config: ConfigOptions = {
//   cloud_name: "fontete",
//   api_key: "865863544799317",
//   api_secret: "7cHdkbZUQQtEimervFBH4Qn6LJc",
// };

// interface RequestWithBody extends Request {
//   body: { file: any };
// }

// cloudinary.v2.config(config);

// exports.upload = async (req: RequestWithBody, res: Response) => {
//   let result = await cloudinary.v2.uploader.upload(req.body.file, {
//     public_id: `${Date.now()}`,
//     resource_type: "auto",
//   });
//   console.log(result);
//   res.json({
//     public_id: result.public_id,
//     url: result.secure_url,
//   });
// };

export const func1 = (req: Request, res: Response) => {
  // const workbook = xlsx.readFile("../../raw.xlsx");
  // const sheetNames = workbook.SheetNames;
  fs.readFile(path.join(__dirname, "../routes/template1.xlsx"), (err, data) => {
    console.log(data)
    const template: XlsxTemplate = new XlsxTemplate(data);
    const sheetNumber: number = 1;
    const values = {
      extractDate: new Date(),
      dates: [
        new Date("2013-06-01"),
        new Date("2013-06-02"),
        new Date("2013-06-03"),
      ],
      people: [
        { name: "John Smith", age: 20 },
        { name: "Bob Johnson", age: 22 },
      ],
    };

    template.substitute(sheetNumber, values);
    const result = template.generate();
    console.log(result)
    res.attachment("generateFile.xlsx");
    res.send(result);
  });
};
