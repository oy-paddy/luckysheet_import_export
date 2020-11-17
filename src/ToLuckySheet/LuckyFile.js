"use strict";
var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
exports.__esModule = true;
exports.LuckyFile = void 0;
var LuckySheet_1 = require("./LuckySheet");
var constant_1 = require("../common/constant");
var ReadXml_1 = require("./ReadXml");
var method_1 = require("../common/method");
var LuckyBase_1 = require("./LuckyBase");
var LuckyImage_1 = require("./LuckyImage");
var LuckyFile = /** @class */ (function (_super) {
    __extends(LuckyFile, _super);
    function LuckyFile(files, fileName) {
        var _this = _super.call(this) || this;
        _this.columnWidthSet = [];
        _this.rowHeightSet = [];
        _this.files = files;
        _this.fileName = fileName;
        _this.readXml = new ReadXml_1.ReadXml(files);
        _this.getSheetNameList();
        _this.sharedStrings = _this.readXml.getElementsByTagName("sst/si", constant_1.sharedStringsFile);
        _this.calcChain = _this.readXml.getElementsByTagName("calcChain/c", constant_1.calcChainFile);
        _this.styles = {};
        _this.styles["cellXfs"] = _this.readXml.getElementsByTagName("cellXfs/xf", constant_1.stylesFile);
        _this.styles["cellStyleXfs"] = _this.readXml.getElementsByTagName("cellStyleXfs/xf", constant_1.stylesFile);
        _this.styles["cellStyles"] = _this.readXml.getElementsByTagName("cellStyles/cellStyle", constant_1.stylesFile);
        _this.styles["fonts"] = _this.readXml.getElementsByTagName("fonts/font", constant_1.stylesFile);
        _this.styles["fills"] = _this.readXml.getElementsByTagName("fills/fill", constant_1.stylesFile);
        _this.styles["borders"] = _this.readXml.getElementsByTagName("borders/border", constant_1.stylesFile);
        _this.styles["clrScheme"] = _this.readXml.getElementsByTagName("a:clrScheme/a:dk1|a:lt1|a:dk2|a:lt2|a:accent1|a:accent2|a:accent3|a:accent4|a:accent5|a:accent6|a:hlink|a:folHlink", constant_1.theme1File);
        _this.styles["indexedColors"] = _this.readXml.getElementsByTagName("colors/indexedColors/rgbColor", constant_1.stylesFile);
        _this.styles["mruColors"] = _this.readXml.getElementsByTagName("colors/mruColors/color", constant_1.stylesFile);
        _this.imageList = new LuckyImage_1.ImageList(files);
        var numfmts = _this.readXml.getElementsByTagName("numFmt/numFmt", constant_1.stylesFile);
        var numFmtDefaultC = constant_1.numFmtDefault;
        for (var i = 0; i < numfmts.length; i++) {
            var attrList = numfmts[i].attributeList;
            var numfmtid = method_1.getXmlAttibute(attrList, "numFmtId", "49");
            var formatcode = method_1.getXmlAttibute(attrList, "formatCode", "@");
            // console.log(numfmtid, formatcode);
            if (!(numfmtid in constant_1.numFmtDefault)) {
                numFmtDefaultC[numfmtid] = formatcode;
            }
        }
        // console.log(JSON.stringify(numFmtDefaultC), numfmts);
        _this.styles["numfmts"] = numFmtDefaultC;
        return _this;
    }
    /**
    * @return All sheet name of workbook
    */
    LuckyFile.prototype.getSheetNameList = function () {
        var workbookRelList = this.readXml.getElementsByTagName("Relationships/Relationship", constant_1.workbookRels);
        if (workbookRelList == null) {
            return;
        }
        var regex = new RegExp("worksheets/[^/]*?.xml");
        var sheetNames = {};
        for (var i = 0; i < workbookRelList.length; i++) {
            var rel = workbookRelList[i], attrList = rel.attributeList;
            var id = attrList["Id"], target = attrList["Target"];
            if (regex.test(target)) {
                sheetNames[id] = "xl/" + target;
            }
        }
        this.sheetNameList = sheetNames;
    };
    /**
    * @param sheetName WorkSheet'name
    * @return sheet file name and path in zip
    */
    LuckyFile.prototype.getSheetFileBysheetId = function (sheetId) {
        // for(let i=0;i<this.sheetNameList.length;i++){
        //     let sheetFileName = this.sheetNameList[i];
        //     if(sheetFileName.indexOf("sheet"+sheetId)>-1){
        //         return sheetFileName;
        //     }
        // }
        return this.sheetNameList[sheetId];
    };
    /**
    * @return workBook information
    */
    LuckyFile.prototype.getWorkBookInfo = function () {
        var Company = this.readXml.getElementsByTagName("Company", constant_1.appFile);
        var AppVersion = this.readXml.getElementsByTagName("AppVersion", constant_1.appFile);
        var creator = this.readXml.getElementsByTagName("dc:creator", constant_1.coreFile);
        var lastModifiedBy = this.readXml.getElementsByTagName("cp:lastModifiedBy", constant_1.coreFile);
        var created = this.readXml.getElementsByTagName("dcterms:created", constant_1.coreFile);
        var modified = this.readXml.getElementsByTagName("dcterms:modified", constant_1.coreFile);
        this.info = new LuckyBase_1.LuckyFileInfo();
        this.info.name = this.fileName;
        this.info.creator = creator.length > 0 ? creator[0].value : "";
        this.info.lastmodifiedby = lastModifiedBy.length > 0 ? lastModifiedBy[0].value : "";
        this.info.createdTime = created.length > 0 ? created[0].value : "";
        this.info.modifiedTime = modified.length > 0 ? modified[0].value : "";
        this.info.company = Company.length > 0 ? Company[0].value : "";
        this.info.appversion = AppVersion.length > 0 ? AppVersion[0].value : "";
    };
    /**
    * @return All sheet , include whole information
    */
    LuckyFile.prototype.getSheetsFull = function (isInitialCell) {
        if (isInitialCell === void 0) { isInitialCell = true; }
        var sheets = this.readXml.getElementsByTagName("sheets/sheet", constant_1.workBookFile);
        var sheetList = {};
        for (var key in sheets) {
            var sheet = sheets[key];
            sheetList[sheet.attributeList.name] = sheet.attributeList["sheetId"];
        }
        this.sheets = [];
        var order = 0;
        for (var key in sheets) {
            var sheet = sheets[key];
            var sheetName = sheet.attributeList.name;
            var sheetId = sheet.attributeList["sheetId"];
            var rid = sheet.attributeList["r:id"];
            var sheetFile = this.getSheetFileBysheetId(rid);
            var drawing = this.readXml.getElementsByTagName("worksheet/drawing", sheetFile), drawingFile = void 0, drawingRelsFile = void 0;
            if (drawing != null && drawing.length > 0) {
                var attrList = drawing[0].attributeList;
                var rid_1 = method_1.getXmlAttibute(attrList, "r:id", null);
                if (rid_1 != null) {
                    drawingFile = this.getDrawingFile(rid_1, sheetFile);
                    drawingRelsFile = this.getDrawingRelsFile(drawingFile);
                }
            }
            if (sheetFile != null) {
                var sheet_1 = new LuckySheet_1.LuckySheet(sheetName, sheetId, order, isInitialCell, {
                    sheetFile: sheetFile,
                    readXml: this.readXml,
                    sheetList: sheetList,
                    styles: this.styles,
                    sharedStrings: this.sharedStrings,
                    calcChain: this.calcChain,
                    imageList: this.imageList,
                    drawingFile: drawingFile,
                    drawingRelsFile: drawingRelsFile
                });
                this.columnWidthSet = [];
                this.rowHeightSet = [];
                this.imagePositionCaculation(sheet_1);
                this.sheets.push(sheet_1);
                order++;
            }
        }
    };
    LuckyFile.prototype.extendArray = function (index, sets, def, hidden, lens) {
        if (index < sets.length) {
            return;
        }
        var startIndex = sets.length, endIndex = index;
        var allGap = 0;
        if (startIndex > 0) {
            allGap = sets[startIndex - 1];
        }
        // else{
        //     sets.push(0);
        // }
        for (var i = startIndex; i <= endIndex; i++) {
            var gap = def, istring = i.toString();
            if (istring in hidden) {
                gap = 0;
            }
            else if (istring in lens) {
                gap = lens[istring];
            }
            allGap += Math.round(gap + 1);
            sets.push(allGap);
        }
    };
    LuckyFile.prototype.imagePositionCaculation = function (sheet) {
        var images = sheet.images, defaultColWidth = sheet.defaultColWidth, defaultRowHeight = sheet.defaultRowHeight;
        var colhidden = {};
        if (sheet.config.colhidden) {
            colhidden = sheet.config.colhidden;
        }
        var columnlen = {};
        if (sheet.config.columnlen) {
            columnlen = sheet.config.columnlen;
        }
        var rowhidden = {};
        if (sheet.config.rowhidden) {
            rowhidden = sheet.config.rowhidden;
        }
        var rowlen = {};
        if (sheet.config.rowlen) {
            rowlen = sheet.config.rowlen;
        }
        for (var key in images) {
            var imageObject = images[key]; //Image, luckyImage
            var fromCol = imageObject.fromCol;
            var fromColOff = imageObject.fromColOff;
            var fromRow = imageObject.fromRow;
            var fromRowOff = imageObject.fromRowOff;
            var toCol = imageObject.toCol;
            var toColOff = imageObject.toColOff;
            var toRow = imageObject.toRow;
            var toRowOff = imageObject.toRowOff;
            var x_n = 0, y_n = 0;
            var cx_n = 0, cy_n = 0;
            if (fromCol >= this.columnWidthSet.length) {
                this.extendArray(fromCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if (fromCol == 0) {
                x_n = 0;
            }
            else {
                x_n = this.columnWidthSet[fromCol - 1];
            }
            x_n = x_n + fromColOff;
            if (fromRow >= this.rowHeightSet.length) {
                this.extendArray(fromRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if (fromRow == 0) {
                y_n = 0;
            }
            else {
                y_n = this.rowHeightSet[fromRow - 1];
            }
            y_n = y_n + fromRowOff;
            if (toCol >= this.columnWidthSet.length) {
                this.extendArray(toCol, this.columnWidthSet, defaultColWidth, colhidden, columnlen);
            }
            if (toCol == 0) {
                cx_n = 0;
            }
            else {
                cx_n = this.columnWidthSet[toCol - 1];
            }
            cx_n = cx_n + toColOff - x_n;
            if (toRow >= this.rowHeightSet.length) {
                this.extendArray(toRow, this.rowHeightSet, defaultRowHeight, rowhidden, rowlen);
            }
            if (toRow == 0) {
                cy_n = 0;
            }
            else {
                cy_n = this.rowHeightSet[toRow - 1];
            }
            cy_n = cy_n + toRowOff - y_n;
            console.log(defaultColWidth, colhidden, columnlen);
            console.log(fromCol, this.columnWidthSet[fromCol], fromColOff);
            console.log(toCol, this.columnWidthSet[toCol], toColOff, JSON.stringify(this.columnWidthSet));
            imageObject.originWidth = cx_n;
            imageObject.originHeight = cy_n;
            imageObject.crop.height = cy_n;
            imageObject.crop.width = cx_n;
            imageObject["default"].height = cy_n;
            imageObject["default"].left = x_n;
            imageObject["default"].top = y_n;
            imageObject["default"].width = cx_n;
        }
        console.log(this.columnWidthSet, this.rowHeightSet);
    };
    /**
    * @return drawing file string
    */
    LuckyFile.prototype.getDrawingFile = function (rid, sheetFile) {
        var sheetRelsPath = "xl/worksheets/_rels/";
        var sheetFileArr = sheetFile.split("/");
        var sheetRelsName = sheetFileArr[sheetFileArr.length - 1];
        var sheetRelsFile = sheetRelsPath + sheetRelsName + ".rels";
        var drawing = this.readXml.getElementsByTagName("Relationships/Relationship", sheetRelsFile);
        if (drawing.length > 0) {
            for (var i = 0; i < drawing.length; i++) {
                var relationship = drawing[i];
                var attrList = relationship.attributeList;
                var relationshipId = method_1.getXmlAttibute(attrList, "Id", null);
                if (relationshipId == rid) {
                    var target = method_1.getXmlAttibute(attrList, "Target", null);
                    if (target != null) {
                        return target.replace(/\.\.\//g, "");
                    }
                }
            }
        }
        return null;
    };
    LuckyFile.prototype.getDrawingRelsFile = function (drawingFile) {
        var drawingRelsPath = "xl/drawings/_rels/";
        var drawingFileArr = drawingFile.split("/");
        var drawingRelsName = drawingFileArr[drawingFileArr.length - 1];
        var drawingRelsFile = drawingRelsPath + drawingRelsName + ".rels";
        return drawingRelsFile;
    };
    /**
    * @return All sheet base information widthout cell and config
    */
    LuckyFile.prototype.getSheetsWithoutCell = function () {
        this.getSheetsFull(false);
    };
    /**
    * @return LuckySheet file json
    */
    LuckyFile.prototype.Parse = function () {
        // let xml = this.readXml;
        // for(let key in this.sheetNameList){
        //     let sheetName=this.sheetNameList[key];
        //     let sheetColumns = xml.getElementsByTagName("row/c/f", sheetName);
        //     console.log(sheetColumns);
        // }
        // return "";
        this.getWorkBookInfo();
        this.getSheetsFull();
        // for(let i=0;i<this.sheets.length;i++){
        //     let sheet = this.sheets[i];
        //     let _borderInfo = sheet.config._borderInfo;
        //     if(_borderInfo==null){
        //         continue;
        //     }
        //     let _borderInfoKeys = Object.keys(_borderInfo);
        //     _borderInfoKeys.sort();
        //     for(let a=0;a<_borderInfoKeys.length;a++){
        //         let key = parseInt(_borderInfoKeys[a]);
        //         let b = _borderInfo[key];
        //         if(b.cells.length==0){
        //             continue;
        //         }
        //         if(sheet.config.borderInfo==null){
        //             sheet.config.borderInfo = [];
        //         }
        //         sheet.config.borderInfo.push(b);
        //     }
        // }
        return this.toJsonString(this);
    };
    LuckyFile.prototype.toJsonString = function (file) {
        var LuckyOutPutFile = new LuckyBase_1.LuckyFileBase();
        LuckyOutPutFile.info = file.info;
        LuckyOutPutFile.sheets = [];
        file.sheets.forEach(function (sheet) {
            var sheetout = new LuckyBase_1.LuckySheetBase();
            //let attrName = ["name","color","config","index","status","order","row","column","luckysheet_select_save","scrollLeft","scrollTop","zoomRatio","showGridLines","defaultColWidth","defaultRowHeight","celldata","chart","isPivotTable","pivotTable","luckysheet_conditionformat_save","freezen","calcChain"];
            if (sheet.name != null) {
                sheetout.name = sheet.name;
            }
            if (sheet.color != null) {
                sheetout.color = sheet.color;
            }
            if (sheet.config != null) {
                sheetout.config = sheet.config;
                // if(sheetout.config._borderInfo!=null){
                //     delete sheetout.config._borderInfo;
                // }
            }
            if (sheet.index != null) {
                sheetout.index = sheet.index;
            }
            if (sheet.status != null) {
                sheetout.status = sheet.status;
            }
            if (sheet.order != null) {
                sheetout.order = sheet.order;
            }
            if (sheet.row != null) {
                sheetout.row = sheet.row;
            }
            if (sheet.column != null) {
                sheetout.column = sheet.column;
            }
            if (sheet.luckysheet_select_save != null) {
                sheetout.luckysheet_select_save = sheet.luckysheet_select_save;
            }
            if (sheet.scrollLeft != null) {
                sheetout.scrollLeft = sheet.scrollLeft;
            }
            if (sheet.scrollTop != null) {
                sheetout.scrollTop = sheet.scrollTop;
            }
            if (sheet.zoomRatio != null) {
                sheetout.zoomRatio = sheet.zoomRatio;
            }
            if (sheet.showGridLines != null) {
                sheetout.showGridLines = sheet.showGridLines;
            }
            if (sheet.defaultColWidth != null) {
                sheetout.defaultColWidth = sheet.defaultColWidth;
            }
            if (sheet.defaultRowHeight != null) {
                sheetout.defaultRowHeight = sheet.defaultRowHeight;
            }
            if (sheet.celldata != null) {
                // sheetout.celldata = sheet.celldata;
                sheetout.celldata = [];
                sheet.celldata.forEach(function (cell) {
                    var cellout = new LuckyBase_1.LuckySheetCelldataBase();
                    cellout.r = cell.r;
                    cellout.c = cell.c;
                    cellout.v = cell.v;
                    sheetout.celldata.push(cellout);
                });
            }
            if (sheet.chart != null) {
                sheetout.chart = sheet.chart;
            }
            if (sheet.isPivotTable != null) {
                sheetout.isPivotTable = sheet.isPivotTable;
            }
            if (sheet.pivotTable != null) {
                sheetout.pivotTable = sheet.pivotTable;
            }
            if (sheet.luckysheet_conditionformat_save != null) {
                sheetout.luckysheet_conditionformat_save = sheet.luckysheet_conditionformat_save;
            }
            if (sheet.freezen != null) {
                sheetout.freezen = sheet.freezen;
            }
            if (sheet.calcChain != null) {
                sheetout.calcChain = sheet.calcChain;
            }
            if (sheet.images != null) {
                sheetout.images = sheet.images;
            }
            LuckyOutPutFile.sheets.push(sheetout);
        });
        return JSON.stringify(LuckyOutPutFile);
    };
    return LuckyFile;
}(LuckyBase_1.LuckyFileBase));
exports.LuckyFile = LuckyFile;
