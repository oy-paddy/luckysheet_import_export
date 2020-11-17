"use strict";
exports.__esModule = true;
exports.LuckyExcel = void 0;
var LuckyFile_1 = require("./ToLuckySheet/LuckyFile");
// import {SecurityDoor,Car} from './content';
var HandleZip_1 = require("./HandleZip");
//demo
function demoHandler() {
    var upload = document.getElementById("Luckyexcel-demo-file");
    var selectADemo = document.getElementById("Luckyexcel-select-demo");
    var downlodDemo = document.getElementById("Luckyexcel-downlod-file");
    var mask = document.getElementById("lucky-mask-demo");
    if (upload) {
        window.onload = function () {
            upload.addEventListener("change", function (evt) {
                var files = evt.target.files;
                if (files == null || files.length == 0) {
                    alert("No files wait for import");
                    return;
                }
                var name = files[0].name;
                var suffixArr = name.split("."), suffix = suffixArr[suffixArr.length - 1];
                if (suffix != "xlsx") {
                    alert("Currently only supports the import of xlsx files");
                    return;
                }
                LuckyExcel.transformExcelToLucky(files[0], function (exportJson, luckysheetfile) {
                    if (exportJson.sheets == null || exportJson.sheets.length == 0) {
                        alert("Failed to read the content of the excel file, currently does not support xls files!");
                        return;
                    }
                    console.log(exportJson, luckysheetfile);
                    window.luckysheet.destroy();
                    window.luckysheet.create({
                        container: 'luckysheet',
                        showinfobar: false,
                        data: exportJson.sheets,
                        title: exportJson.info.name,
                        userInfo: exportJson.info.name.creator
                    });
                });
            });
            selectADemo.addEventListener("change", function (evt) {
                var obj = selectADemo;
                var index = obj.selectedIndex;
                var value = obj.options[index].value;
                var name = obj.options[index].innerHTML;
                if (value == "") {
                    return;
                }
                mask.style.display = "flex";
                LuckyExcel.transformExcelToLuckyByUrl(value, name, function (exportJson, luckysheetfile) {
                    if (exportJson.sheets == null || exportJson.sheets.length == 0) {
                        alert("Failed to read the content of the excel file, currently does not support xls files!");
                        return;
                    }
                    console.log(exportJson, luckysheetfile);
                    mask.style.display = "none";
                    window.luckysheet.destroy();
                    window.luckysheet.create({
                        container: 'luckysheet',
                        showinfobar: false,
                        data: exportJson.sheets,
                        title: exportJson.info.name,
                        userInfo: exportJson.info.name.creator
                    });
                });
            });
            downlodDemo.addEventListener("click", function (evt) {
                var obj = selectADemo;
                var index = obj.selectedIndex;
                var value = obj.options[index].value;
                if (value.length == 0) {
                    alert("Please select a demo file");
                    return;
                }
                var elemIF = document.getElementById("Lucky-download-frame");
                if (elemIF == null) {
                    elemIF = document.createElement("iframe");
                    elemIF.style.display = "none";
                    elemIF.id = "Lucky-download-frame";
                    document.body.appendChild(elemIF);
                }
                elemIF.src = value;
                // elemIF.parentNode.removeChild(elemIF);
            });
        };
    }
    LuckyExcel.transformExcelToLuckyByUrl("http://106.55.249.40/publicApi/downloadpublic?fileDir=&fileName=Q4-022-02%E6%88%90%E5%93%81%E5%B7%A1%E6%AA%A2%E9%BB%9E%E6%AA%A2%E8%A1%A8_new20200720_15_51_38.xlsx", "Q4-022-02成品巡檢點檢表_new20200720_15_51_38.xlsx", function (exportJson, luckysheetfile) {
        if (exportJson.sheets == null || exportJson.sheets.length == 0) {
            alert("Failed to read the content of the excel file, currently does not support xls files!");
            return;
        }
        console.log(exportJson, luckysheetfile);
        mask.style.display = "none";
        window.luckysheet.destroy();
        window.luckysheet.create({
            container: 'luckysheet',
            showinfobar: false,
            data: exportJson.sheets,
            title: exportJson.info.name,
            userInfo: exportJson.info.name.creator
        });
    });
}
demoHandler();
// api
var LuckyExcel = /** @class */ (function () {
    function LuckyExcel() {
    }
    LuckyExcel.transformExcelToLucky = function (excelFile, callBack) {
        var handleZip = new HandleZip_1.HandleZip(excelFile);
        handleZip.unzipFile(function (files) {
            var luckyFile = new LuckyFile_1.LuckyFile(files, excelFile.name);
            var luckysheetfile = luckyFile.Parse();
            var exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        }, function (err) {
            console.error(err);
        });
    };
    LuckyExcel.transformExcelToLuckyByUrl = function (url, name, callBack) {
        var handleZip = new HandleZip_1.HandleZip();
        handleZip.unzipFileByUrl(url, function (files) {
            var luckyFile = new LuckyFile_1.LuckyFile(files, name);
            var luckysheetfile = luckyFile.Parse();
            var exportJson = JSON.parse(luckysheetfile);
            if (callBack != undefined) {
                callBack(exportJson, luckysheetfile);
            }
        }, function (err) {
            console.error(err);
        });
    };
    LuckyExcel.transformLuckyToExcel = function (LuckyFile, callBack) {
    };
    return LuckyExcel;
}());
exports.LuckyExcel = LuckyExcel;
