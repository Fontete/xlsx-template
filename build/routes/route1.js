"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.router1 = void 0;
var express_1 = require("express");
var xlsx_template_1 = __importDefault(require("xlsx-template"));
var fs_1 = __importDefault(require("fs"));
var path_1 = __importDefault(require("path"));
var xlsx_1 = __importDefault(require("xlsx"));
// import {func1} from '../controllers/controller1'
var router1 = express_1.Router();
exports.router1 = router1;
// router1.get("/", func1);
router1.post("/", function (req, res) {
    var workbook = xlsx_1.default.readFile("./raw.xlsx");
    var sheet_name_list = workbook.SheetNames;
    // type inputData = Object[];
    var sheet1 = workbook.Sheets[sheet_name_list[0]];
    var rawData = xlsx_1.default.utils.sheet_to_json(sheet1, {
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
    fs_1.default.readFile(path_1.default.join(__dirname, "../../template.xlsx"), function (err, data) {
        var template = new xlsx_template_1.default(data);
        var sheetNumber = 1;
        var values = {
            // extractDate: new Date(),
            // dates: [
            //   new Date("2013-06-01"),
            //   new Date("2013-06-02"),
            //   new Date("2013-06-03"),
            // ],
            rawData: rawData,
        };
        console.log(values);
        template.substitute(sheetNumber, values);
        var result = template.generate({ type: "nodebuffer" });
        res.attachment("generateFile.xlsx");
        // xlsx.read(result, { type: "buffer" }).SheetNames.forEach((s) => {});
        // console.log("result", xlsx.read(result, { type: "buffer" }).Sheets.Sheet1);
        res.send(result);
    });
});
