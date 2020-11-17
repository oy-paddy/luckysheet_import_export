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
exports.LuckySheet = void 0;
var LuckyCell_1 = require("./LuckyCell");
var method_1 = require("../common/method");
var ReadXml_1 = require("./ReadXml");
var LuckyBase_1 = require("./LuckyBase");
var LuckySheet = /** @class */ (function (_super) {
    __extends(LuckySheet, _super);
    function LuckySheet(sheetName, sheetId, sheetOrder, isInitialCell, allFileOption) {
        if (isInitialCell === void 0) { isInitialCell = false; }
        var _this = 
        //Private
        _super.call(this) || this;
        _this.isInitialCell = isInitialCell;
        _this.readXml = allFileOption.readXml;
        _this.sheetFile = allFileOption.sheetFile;
        _this.styles = allFileOption.styles;
        _this.sharedStrings = allFileOption.sharedStrings;
        _this.calcChainEles = allFileOption.calcChain;
        _this.sheetList = allFileOption.sheetList;
        _this.imageList = allFileOption.imageList;
        //Output
        _this.name = sheetName;
        _this.index = sheetId;
        _this.order = sheetOrder.toString();
        _this.config = new LuckyBase_1.LuckyConfig();
        _this.celldata = [];
        _this.mergeCells = _this.readXml.getElementsByTagName("mergeCells/mergeCell", _this.sheetFile);
        var clrScheme = _this.styles["clrScheme"];
        var sheetView = _this.readXml.getElementsByTagName("sheetViews/sheetView", _this.sheetFile);
        var showGridLines = "1", tabSelected = "0", zoomScale = "100", activeCell = "A1";
        if (sheetView.length > 0) {
            var attrList = sheetView[0].attributeList;
            showGridLines = method_1.getXmlAttibute(attrList, "showGridLines", "1");
            tabSelected = method_1.getXmlAttibute(attrList, "tabSelected", "0");
            zoomScale = method_1.getXmlAttibute(attrList, "zoomScale", "100");
            // let colorId = getXmlAttibute(attrList, "colorId", "0");
            var selections = sheetView[0].getInnerElements("selection");
            if (selections != null && selections.length > 0) {
                activeCell = method_1.getXmlAttibute(selections[0].attributeList, "activeCell", "A1");
                var range = method_1.getcellrange(activeCell, _this.sheetList, sheetId);
                _this.luckysheet_select_save = [];
                _this.luckysheet_select_save.push(range);
            }
        }
        _this.showGridLines = showGridLines;
        _this.status = tabSelected;
        _this.zoomRatio = parseInt(zoomScale) / 100;
        var tabColors = _this.readXml.getElementsByTagName("sheetPr/tabColor", _this.sheetFile);
        if (tabColors != null && tabColors.length > 0) {
            var tabColor = tabColors[0], attrList = tabColor.attributeList;
            // if(attrList.rgb!=null){
            var tc = ReadXml_1.getColor(tabColor, _this.styles, "b");
            _this.color = tc;
            // }
        }
        var sheetFormatPr = _this.readXml.getElementsByTagName("sheetFormatPr", _this.sheetFile);
        var defaultColWidth, defaultRowHeight;
        if (sheetFormatPr.length > 0) {
            var attrList = sheetFormatPr[0].attributeList;
            defaultColWidth = method_1.getXmlAttibute(attrList, "defaultColWidth", "9.21");
            defaultRowHeight = method_1.getXmlAttibute(attrList, "defaultRowHeight", "19");
        }
        _this.defaultColWidth = method_1.getColumnWidthPixel(parseFloat(defaultColWidth));
        _this.defaultRowHeight = method_1.getRowHeightPixel(parseFloat(defaultRowHeight));
        _this.generateConfigColumnLenAndHidden();
        _this.generateConfigRowLenAndHiddenAddCell();
        if (_this.formulaRefList != null) {
            for (var key in _this.formulaRefList) {
                var funclist = _this.formulaRefList[key];
                var mainFunc = funclist["mainRef"], mainCellValue = mainFunc.cellValue;
                var formulaTxt = mainFunc.fv;
                var mainR = mainCellValue.r, mainC = mainCellValue.c;
                // let refRange = getcellrange(ref);
                for (var name_1 in funclist) {
                    if (name_1 == "mainRef") {
                        continue;
                    }
                    var funcValue = funclist[name_1], cellValue = funcValue.cellValue;
                    if (cellValue == null) {
                        continue;
                    }
                    var r = cellValue.r, c = cellValue.c;
                    var func = formulaTxt;
                    var offsetRow = r - mainR, offsetCol = c - mainC;
                    if (offsetRow > 0) {
                        func = "=" + method_1.fromulaRef.functionCopy(func, "down", offsetRow);
                    }
                    else if (offsetRow < 0) {
                        func = "=" + method_1.fromulaRef.functionCopy(func, "up", Math.abs(offsetRow));
                    }
                    if (offsetCol > 0) {
                        func = "=" + method_1.fromulaRef.functionCopy(func, "right", offsetCol);
                    }
                    else if (offsetCol < 0) {
                        func = "=" + method_1.fromulaRef.functionCopy(func, "left", Math.abs(offsetCol));
                    }
                    // console.log(offsetRow, offsetCol, func);
                    cellValue.v.f = func;
                }
            }
        }
        if (_this.calcChain == null) {
            _this.calcChain = [];
        }
        for (var c = 0; c < _this.calcChainEles.length; c++) {
            var calcChainEle = _this.calcChainEles[c], attrList = calcChainEle.attributeList;
            if (attrList.i != sheetId) {
                continue;
            }
            var r = attrList.r, i = attrList.i, l = attrList.l, s = attrList.s, a = attrList.a, t = attrList.t;
            var range = method_1.getcellrange(r);
            var chain = new LuckyBase_1.LuckysheetCalcChain();
            chain.r = range.row[0];
            chain.c = range.column[0];
            chain.index = _this.index;
            _this.calcChain.push(chain);
        }
        if (_this.mergeCells != null) {
            for (var i = 0; i < _this.mergeCells.length; i++) {
                var merge = _this.mergeCells[i], attrList = merge.attributeList;
                var ref = attrList.ref;
                if (ref == null) {
                    continue;
                }
                var range = method_1.getcellrange(ref, _this.sheetList, sheetId);
                var mergeValue = new LuckyBase_1.LuckySheetConfigMerge();
                mergeValue.r = range.row[0];
                mergeValue.c = range.column[0];
                mergeValue.rs = range.row[1] - range.row[0] + 1;
                mergeValue.cs = range.column[1] - range.column[0] + 1;
                if (_this.config.merge == null) {
                    _this.config.merge = {};
                }
                _this.config.merge[range.row[0] + "_" + range.column[0]] = mergeValue;
            }
        }
        var drawingFile = allFileOption.drawingFile, drawingRelsFile = allFileOption.drawingRelsFile;
        if (drawingFile != null && drawingRelsFile != null) {
            var twoCellAnchors = _this.readXml.getElementsByTagName("xdr:twoCellAnchor", drawingFile);
            if (twoCellAnchors != null && twoCellAnchors.length > 0) {
                for (var i = 0; i < twoCellAnchors.length; i++) {
                    var twoCellAnchor = twoCellAnchors[i];
                    var editAs = method_1.getXmlAttibute(twoCellAnchor.attributeList, "editAs", "twoCell");
                    var xdrFroms = twoCellAnchor.getInnerElements("xdr:from"), xdrTos = twoCellAnchor.getInnerElements("xdr:to");
                    var xdr_blipfills = twoCellAnchor.getInnerElements("a:blip");
                    if (xdrFroms != null && xdr_blipfills != null && xdrFroms.length > 0 && xdr_blipfills.length > 0) {
                        var xdrFrom = xdrFroms[0], xdrTo = xdrTos[0], xdr_blipfill = xdr_blipfills[0];
                        var rembed = method_1.getXmlAttibute(xdr_blipfill.attributeList, "r:embed", null);
                        var imageObject = _this.getBase64ByRid(rembed, drawingRelsFile);
                        // let aoff = xdr_xfrm.getInnerElements("a:off"), aext = xdr_xfrm.getInnerElements("a:ext");
                        // if(aoff!=null && aext!=null && aoff.length>0 && aext.length>0){
                        //     let aoffAttribute = aoff[0].attributeList, aextAttribute = aext[0].attributeList;
                        //     let x = getXmlAttibute(aoffAttribute, "x", null);
                        //     let y = getXmlAttibute(aoffAttribute, "y", null);
                        //     let cx = getXmlAttibute(aextAttribute, "cx", null);
                        //     let cy = getXmlAttibute(aextAttribute, "cy", null);
                        //     if(x!=null && y!=null && cx!=null && cy!=null && imageObject !=null){
                        // let x_n = getPxByEMUs(parseInt(x), "c"),y_n = getPxByEMUs(parseInt(y));
                        // let cx_n = getPxByEMUs(parseInt(cx), "c"),cy_n = getPxByEMUs(parseInt(cy));
                        var x_n = 0, y_n = 0;
                        var cx_n = 0, cy_n = 0;
                        imageObject.fromCol = _this.getXdrValue(xdrFrom.getInnerElements("xdr:col"));
                        imageObject.fromColOff = method_1.getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:colOff")));
                        imageObject.fromRow = _this.getXdrValue(xdrFrom.getInnerElements("xdr:row"));
                        imageObject.fromRowOff = method_1.getPxByEMUs(_this.getXdrValue(xdrFrom.getInnerElements("xdr:rowOff")));
                        imageObject.toCol = _this.getXdrValue(xdrTo.getInnerElements("xdr:col"));
                        imageObject.toColOff = method_1.getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:colOff")));
                        imageObject.toRow = _this.getXdrValue(xdrTo.getInnerElements("xdr:row"));
                        imageObject.toRowOff = method_1.getPxByEMUs(_this.getXdrValue(xdrTo.getInnerElements("xdr:rowOff")));
                        imageObject.originWidth = cx_n;
                        imageObject.originHeight = cy_n;
                        if (editAs == "absolute") {
                            imageObject.type = "3";
                        }
                        else if (editAs == "oneCell") {
                            imageObject.type = "2";
                        }
                        else {
                            imageObject.type = "1";
                        }
                        imageObject.isFixedPos = false;
                        imageObject.fixedLeft = 0;
                        imageObject.fixedTop = 0;
                        var imageBorder = {
                            color: "#000",
                            radius: 0,
                            style: "solid",
                            width: 0
                        };
                        imageObject.border = imageBorder;
                        var imageCrop = {
                            height: cy_n,
                            offsetLeft: 0,
                            offsetTop: 0,
                            width: cx_n
                        };
                        imageObject.crop = imageCrop;
                        var imageDefault = {
                            height: cy_n,
                            left: x_n,
                            top: y_n,
                            width: cx_n
                        };
                        imageObject["default"] = imageDefault;
                        if (_this.images == null) {
                            _this.images = {};
                        }
                        _this.images[method_1.generateRandomIndex("image")] = imageObject;
                        //     }
                        // }
                    }
                }
            }
        }
        return _this;
    }
    LuckySheet.prototype.getXdrValue = function (ele) {
        if (ele == null || ele.length == 0) {
            return null;
        }
        return parseInt(ele[0].value);
    };
    LuckySheet.prototype.getBase64ByRid = function (rid, drawingRelsFile) {
        var Relationships = this.readXml.getElementsByTagName("Relationships/Relationship", drawingRelsFile);
        if (Relationships != null && Relationships.length > 0) {
            for (var i = 0; i < Relationships.length; i++) {
                var Relationship = Relationships[i];
                var attrList = Relationship.attributeList;
                var Id = method_1.getXmlAttibute(attrList, "Id", null);
                var src = method_1.getXmlAttibute(attrList, "Target", null);
                if (Id == rid) {
                    src = src.replace(/\.\.\//g, "");
                    src = "xl/" + src;
                    var imgage = this.imageList.getImageByName(src);
                    return imgage;
                }
            }
        }
        return null;
    };
    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    LuckySheet.prototype.generateConfigColumnLenAndHidden = function () {
        var cols = this.readXml.getElementsByTagName("cols/col", this.sheetFile);
        for (var i = 0; i < cols.length; i++) {
            var col = cols[i], attrList = col.attributeList;
            var min = method_1.getXmlAttibute(attrList, "min", null);
            var max = method_1.getXmlAttibute(attrList, "max", null);
            var width = method_1.getXmlAttibute(attrList, "width", null);
            var hidden = method_1.getXmlAttibute(attrList, "hidden", null);
            var customWidth = method_1.getXmlAttibute(attrList, "customWidth", null);
            if (min == null || max == null) {
                continue;
            }
            var minNum = parseInt(min) - 1, maxNum = parseInt(max) - 1, widthNum = parseFloat(width);
            for (var m = minNum; m <= maxNum; m++) {
                if (width != null) {
                    if (this.config.columnlen == null) {
                        this.config.columnlen = {};
                    }
                    this.config.columnlen[m] = method_1.getColumnWidthPixel(widthNum);
                }
                if (hidden == "1") {
                    if (this.config.colhidden == null) {
                        this.config.colhidden = {};
                    }
                    this.config.colhidden[m] = 0;
                    if (this.config.columnlen) {
                        delete this.config.columnlen[m];
                    }
                }
                if (customWidth != null) {
                    if (this.config.customWidth == null) {
                        this.config.customWidth = {};
                    }
                    this.config.customWidth[m] = 1;
                }
            }
        }
    };
    /**
    * @desc This will convert cols/col to luckysheet config of column'width
    */
    LuckySheet.prototype.generateConfigRowLenAndHiddenAddCell = function () {
        var rows = this.readXml.getElementsByTagName("sheetData/row", this.sheetFile);
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i], attrList = row.attributeList;
            var rowNo = method_1.getXmlAttibute(attrList, "r", null);
            var height = method_1.getXmlAttibute(attrList, "ht", null);
            var hidden = method_1.getXmlAttibute(attrList, "hidden", null);
            var customHeight = method_1.getXmlAttibute(attrList, "customHeight", null);
            if (rowNo == null) {
                continue;
            }
            var rowNoNum = parseInt(rowNo) - 1;
            if (height != null) {
                var heightNum = parseFloat(height);
                if (this.config.rowlen == null) {
                    this.config.rowlen = {};
                }
                this.config.rowlen[rowNoNum] = method_1.getRowHeightPixel(heightNum);
            }
            if (hidden == "1") {
                if (this.config.rowhidden == null) {
                    this.config.rowhidden = {};
                }
                this.config.rowhidden[rowNoNum] = 0;
                if (this.config.rowlen) {
                    delete this.config.rowlen[rowNoNum];
                }
            }
            if (customHeight != null) {
                if (this.config.customHeight == null) {
                    this.config.customHeight = {};
                }
                this.config.customHeight[rowNoNum] = 1;
            }
            if (this.isInitialCell) {
                var cells = row.getInnerElements("c");
                for (var key in cells) {
                    var cell = cells[key];
                    var cellValue = new LuckyCell_1.LuckySheetCelldata(cell, this.styles, this.sharedStrings, this.mergeCells, this.sheetFile, this.readXml);
                    if (cellValue._borderObject != null) {
                        if (this.config.borderInfo == null) {
                            this.config.borderInfo = [];
                        }
                        this.config.borderInfo.push(cellValue._borderObject);
                        delete cellValue._borderObject;
                    }
                    // let borderId = cellValue._borderId;
                    // if(borderId!=null){
                    //     let borders = this.styles["borders"] as Element[];
                    //     if(this.config._borderInfo==null){
                    //         this.config._borderInfo = {};
                    //     }
                    //     if( borderId in this.config._borderInfo){
                    //         this.config._borderInfo[borderId].cells.push(cellValue.r + "_" + cellValue.c);
                    //     }
                    //     else{
                    //         let border = borders[borderId];
                    //         let borderObject = new LuckySheetborderInfoCellForImp();
                    //         borderObject.rangeType = "cellGroup";
                    //         borderObject.cells = [];
                    //         let borderCellValue = new LuckySheetborderInfoCellValue();
                    //         let lefts = border.getInnerElements("left");
                    //         let rights = border.getInnerElements("right");
                    //         let tops = border.getInnerElements("top");
                    //         let bottoms = border.getInnerElements("bottom");
                    //         let diagonals = border.getInnerElements("diagonal");
                    //         let left = this.getBorderInfo(lefts);
                    //         let right = this.getBorderInfo(rights);
                    //         let top = this.getBorderInfo(tops);
                    //         let bottom = this.getBorderInfo(bottoms);
                    //         let diagonal = this.getBorderInfo(diagonals);
                    //         let isAdd = false;
                    //         if(left!=null && left.color!=null){
                    //             borderCellValue.l = left;
                    //             isAdd = true;
                    //         }
                    //         if(right!=null && right.color!=null){
                    //             borderCellValue.r = right;
                    //             isAdd = true;
                    //         }
                    //         if(top!=null && top.color!=null){
                    //             borderCellValue.t = top;
                    //             isAdd = true;
                    //         }
                    //         if(bottom!=null && bottom.color!=null){
                    //             borderCellValue.b = bottom;
                    //             isAdd = true;
                    //         }
                    //         if(isAdd){
                    //             borderObject.value = borderCellValue;
                    //             this.config._borderInfo[borderId] = borderObject;
                    //         }
                    //     }
                    // }
                    if (cellValue._formulaType == "shared") {
                        if (this.formulaRefList == null) {
                            this.formulaRefList = {};
                        }
                        if (this.formulaRefList[cellValue._formulaSi] == null) {
                            this.formulaRefList[cellValue._formulaSi] = {};
                        }
                        var fv = void 0;
                        if (cellValue.v != null) {
                            fv = cellValue.v.f;
                        }
                        var refValue = {
                            t: cellValue._formulaType,
                            ref: cellValue._fomulaRef,
                            si: cellValue._formulaSi,
                            fv: fv,
                            cellValue: cellValue
                        };
                        if (cellValue._fomulaRef != null) {
                            this.formulaRefList[cellValue._formulaSi]["mainRef"] = refValue;
                        }
                        else {
                            this.formulaRefList[cellValue._formulaSi][cellValue.r + "_" + cellValue.c] = refValue;
                        }
                        // console.log(refValue, this.formulaRefList);
                    }
                    this.celldata.push(cellValue);
                }
            }
        }
    };
    return LuckySheet;
}(LuckyBase_1.LuckySheetBase));
exports.LuckySheet = LuckySheet;
