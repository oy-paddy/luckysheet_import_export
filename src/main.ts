import {LuckyFile} from "./ToLuckySheet/LuckyFile";
// import {SecurityDoor,Car} from './content';

import {HandleZip} from './HandleZip';

import {IuploadfileList} from "./ICommon";
import {LuckySheet} from "./ToLuckySheet/LuckySheet";


function demoHandler2(): void {
    //var upload = document.getElementById("Luckyexcel-upload");

    window.onload = () => {
        var value_ = document.getElementById("viewpath")as HTMLInputElement;//文件地址
        var name_ = document.getElementById("viewname")as HTMLInputElement;//文件名


        console.log("value_ = " + value_.value + ";name_ = " + name_.value);

        LuckyExcel.transformExcelToLuckyByUrl(value_.value, name_.value, function (exportJson: any, luckysheetfile: string) {

            if (exportJson.sheets == null || exportJson.sheets.length == 0) {
                alert("Failed to read the content of the excel file, currently does not support xls files!");
                return;
            }
            console.log(exportJson, luckysheetfile);
            window.luckysheet.destroy();

            window.luckysheet.create({
                container: 'luckysheet', //luckysheet is the container id
                showinfobar: true,
                data: exportJson.sheets,
                title: name_.value,
                userInfo: exportJson.info.name.creator,
                lang: 'zh', // 设定表格语言
                functionButton: '<button id="uploadLuckysheet" class="btn btn-primary" onclick="uploadExcelData(window.luckysheet)"  style="padding:3px 6px;font-size: 12px;margin-right: 10px;">保存</button> <button id="" class="btn btn-primary btn-danger" onclick="downExcelData(window.luckysheet)" style=" padding:3px 6px; font-size: 12px; margin-right: 85px;" >下载</button>',

            });



        });

    }
}

demoHandler2();


// api
export class LuckyExcel {
    static transformExcelToLucky(excelFile: File, callBack?: (files: IuploadfileList, fs?: string) => void) {
        let handleZip: HandleZip = new HandleZip(excelFile);
        handleZip.unzipFile(function (files: IuploadfileList) {
                let luckyFile = new LuckyFile(files, excelFile.name);
                let luckysheetfile = luckyFile.Parse();
                let exportJson = JSON.parse(luckysheetfile);
                if (callBack != undefined) {
                    callBack(exportJson, luckysheetfile);
                }

            },
            function (err: Error) {
                console.error(err);
            });
    }

    static transformExcelToLuckyByUrl(url: string, name: string, callBack?: (files: IuploadfileList, fs?: string) => void) {
        let handleZip: HandleZip = new HandleZip();
        handleZip.unzipFileByUrl(url, function (files: IuploadfileList) {
                let luckyFile = new LuckyFile(files, name);
                let luckysheetfile = luckyFile.Parse();
                let exportJson = JSON.parse(luckysheetfile);
                if (callBack != undefined) {
                    callBack(exportJson, luckysheetfile);
                }
            },
            function (err: Error) {
                console.error(err);
            });
    }

    static transformLuckyToExcel(LuckyFile: any, callBack?: (files: string) => void) {

    }
}



