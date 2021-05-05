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
// import xlsxChart from "xlsx-chart";
// import {func1} from '../controllers/controller1'
var router1 = express_1.Router();
exports.router1 = router1;
// router1.get("/", func1);
router1.post("/", function (req, res) {
    var workbook = xlsx_1.default.readFile("./template.xlsx");
    var sheet_name_list = workbook.SheetNames;
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
    var sheet1 = workbook.Sheets[sheet_name_list[0]];
    var sheet2 = workbook.Sheets[sheet_name_list[1]];
    var sheet3 = workbook.Sheets[sheet_name_list[2]];
    var sheet4 = workbook.Sheets[sheet_name_list[3]];
    var sheet5 = workbook.Sheets[sheet_name_list[4]];
    var sheet6 = workbook.Sheets[sheet_name_list[5]];
    var sheet7 = workbook.Sheets[sheet_name_list[6]];
    var sheet8 = workbook.Sheets[sheet_name_list[7]];
    var sheet9 = workbook.Sheets[sheet_name_list[8]];
    var sheet10 = workbook.Sheets[sheet_name_list[9]];
    var rawData = xlsx_1.default.utils.sheet_to_json(sheet3, {
        header: "A",
        raw: true,
        blankrows: false,
        // defval: null,
    });
    var newData = [];
    for (var i = 0; i < rawData.length; i++) {
        if ((rawData[i].F && rawData[i].G) !== (null || undefined)) {
            var obj = { F: rawData[i].F, G: rawData[i].G };
            newData.push(obj);
        }
    }
    console.log(newData);
    var rawData2 = xlsx_1.default.utils.sheet_to_json(sheet4, {
        header: "A",
        raw: true,
        blankrows: false,
        defval: null,
    });
    var rawData3 = xlsx_1.default.utils.sheet_to_json(sheet5, {
        header: "A",
        raw: true,
        blankrows: false,
        defval: null,
    });
    //console.log("raw", rawData);
    fs_1.default.readFile(path_1.default.join(__dirname, "../../template.xlsx"), function (err, data) {
        var template = new xlsx_template_1.default(data);
        var sheetNumber = 2;
        var values = {
            rawData: rawData,
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
        var result = template.generate({ type: "nodebuffer" });
        res.attachment("generateFile.xlsx");
        // xlsx.read(result, { type: "buffer" }).SheetNames.forEach((s) => {});
        // console.log("result", xlsx.read(result, { typC: "buffer" }).Sheets);
        res.send(result);
    });
});
