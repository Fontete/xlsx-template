"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.func1 = void 0;
var xlsx_template_1 = __importDefault(require("xlsx-template"));
var fs_1 = __importDefault(require("fs"));
var path_1 = __importDefault(require("path"));
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
var func1 = function (req, res) {
    // const workbook = xlsx.readFile("../../raw.xlsx");
    // const sheetNames = workbook.SheetNames;
    fs_1.default.readFile(path_1.default.join(__dirname, "../routes/template1.xlsx"), function (err, data) {
        console.log(data);
        var template = new xlsx_template_1.default(data);
        var sheetNumber = 1;
        var values = {
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
        var result = template.generate();
        console.log(result);
        res.attachment("generateFile.xlsx");
        res.send(result);
    });
};
exports.func1 = func1;
