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
exports.LuckySheetCelldata = void 0;
var ReadXml_1 = require("./ReadXml");
var method_1 = require("../common/method");
var constant_1 = require("../common/constant");
var LuckyBase_1 = require("./LuckyBase");
var LuckySheetCelldata = /** @class */ (function (_super) {
    __extends(LuckySheetCelldata, _super);
    function LuckySheetCelldata(cell, styles, sharedStrings, mergeCells, sheetFile, ReadXml) {
        var _this = 
        //Private
        _super.call(this) || this;
        _this.cell = cell;
        _this.sheetFile = sheetFile;
        _this.styles = styles;
        _this.sharedStrings = sharedStrings;
        _this.readXml = ReadXml;
        _this.mergeCells = mergeCells;
        var attrList = cell.attributeList;
        var r = attrList.r, s = attrList.s, t = attrList.t;
        var range = method_1.getcellrange(r);
        _this.r = range.row[0];
        _this.c = range.column[0];
        _this.v = _this.generateValue(s, t);
        return _this;
    }
    /**
    * @param s Style index ,start 1
    * @param t Cell type, Optional value is ST_CellType, it's found at constat.ts
    */
    LuckySheetCelldata.prototype.generateValue = function (s, t) {
        var _this = this;
        var v = this.cell.getInnerElements("v");
        var f = this.cell.getInnerElements("f");
        if (v == null) {
            v = this.cell.getInnerElements("t");
        }
        var cellXfs = this.styles["cellXfs"];
        var cellStyleXfs = this.styles["cellStyleXfs"];
        var cellStyles = this.styles["cellStyles"];
        var fonts = this.styles["fonts"];
        var fills = this.styles["fills"];
        var borders = this.styles["borders"];
        var numfmts = this.styles["numfmts"];
        var clrScheme = this.styles["clrScheme"];
        var sharedStrings = this.sharedStrings;
        var cellValue = new LuckyBase_1.LuckySheetCelldataValue();
        if (f != null) {
            var formula = f[0], attrList = formula.attributeList;
            var t_1 = attrList.t, ref = attrList.ref, si = attrList.si;
            var formulaValue = f[0].value;
            if (t_1 == "shared") {
                this._fomulaRef = ref;
                this._formulaType = t_1;
                this._formulaSi = si;
            }
            // console.log(ref, t, si);
            if (ref != null || (formulaValue != null && formulaValue.length > 0)) {
                formulaValue = method_1.escapeCharacter(formulaValue);
                cellValue.f = "=" + formulaValue;
            }
        }
        var familyFont = null;
        var quotePrefix;
        if (s != null) {
            var sNum = parseInt(s);
            var cellXf = cellXfs[sNum];
            var xfId = cellXf.attributeList.xfId;
            var numFmtId = void 0, fontId = void 0, fillId = void 0, borderId = void 0;
            var horizontal = void 0, vertical = void 0, wrapText = void 0, textRotation = void 0, shrinkToFit = void 0, indent = void 0, applyProtection = void 0;
            if (xfId != null) {
                var cellStyleXf = cellStyleXfs[parseInt(xfId)];
                var attrList = cellStyleXf.attributeList;
                var applyNumberFormat_1 = attrList.applyNumberFormat;
                var applyFont_1 = attrList.applyFont;
                var applyFill_1 = attrList.applyFill;
                var applyBorder_1 = attrList.applyBorder;
                var applyAlignment_1 = attrList.applyAlignment;
                // let applyProtection = attrList.applyProtection;
                applyProtection = attrList.applyProtection;
                quotePrefix = attrList.quotePrefix;
                if (applyNumberFormat_1 != "0" && attrList.numFmtId != null) {
                    // if(attrList.numFmtId!="0"){
                    numFmtId = attrList.numFmtId;
                    // }
                }
                if (applyFont_1 != "0" && attrList.fontId != null) {
                    fontId = attrList.fontId;
                }
                if (applyFill_1 != "0" && attrList.fillId != null) {
                    fillId = attrList.fillId;
                }
                if (applyBorder_1 != "0" && attrList.borderId != null) {
                    borderId = attrList.borderId;
                }
                if (applyAlignment_1 != null && applyAlignment_1 != "0") {
                    var alignment = cellStyleXf.getInnerElements("alignment");
                    if (alignment != null) {
                        var attrList_1 = alignment[0].attributeList;
                        if (attrList_1.horizontal != null) {
                            horizontal = attrList_1.horizontal;
                        }
                        if (attrList_1.vertical != null) {
                            vertical = attrList_1.vertical;
                        }
                        if (attrList_1.wrapText != null) {
                            wrapText = attrList_1.wrapText;
                        }
                        if (attrList_1.textRotation != null) {
                            textRotation = attrList_1.textRotation;
                        }
                        if (attrList_1.shrinkToFit != null) {
                            shrinkToFit = attrList_1.shrinkToFit;
                        }
                        if (attrList_1.indent != null) {
                            indent = attrList_1.indent;
                        }
                    }
                }
            }
            var applyNumberFormat = cellXf.attributeList.applyNumberFormat;
            var applyFont = cellXf.attributeList.applyFont;
            var applyFill = cellXf.attributeList.applyFill;
            var applyBorder = cellXf.attributeList.applyBorder;
            var applyAlignment = cellXf.attributeList.applyAlignment;
            if (cellXf.attributeList.applyProtection != null) {
                applyProtection = cellXf.attributeList.applyProtection;
            }
            if (cellXf.attributeList.quotePrefix != null) {
                quotePrefix = cellXf.attributeList.quotePrefix;
            }
            if (applyNumberFormat != "0" && cellXf.attributeList.numFmtId != null) {
                numFmtId = cellXf.attributeList.numFmtId;
            }
            if (applyFont != "0") {
                fontId = cellXf.attributeList.fontId;
            }
            if (applyFill != "0") {
                fillId = cellXf.attributeList.fillId;
            }
            if (applyBorder != "0") {
                borderId = cellXf.attributeList.borderId;
            }
            if (applyAlignment != "0") {
                var alignment = cellXf.getInnerElements("alignment");
                if (alignment != null && alignment.length > 0) {
                    var attrList = alignment[0].attributeList;
                    if (attrList.horizontal != null) {
                        horizontal = attrList.horizontal;
                    }
                    if (attrList.vertical != null) {
                        vertical = attrList.vertical;
                    }
                    if (attrList.wrapText != null) {
                        wrapText = attrList.wrapText;
                    }
                    if (attrList.textRotation != null) {
                        textRotation = attrList.textRotation;
                    }
                    if (attrList.shrinkToFit != null) {
                        shrinkToFit = attrList.shrinkToFit;
                    }
                    if (attrList.indent != null) {
                        indent = attrList.indent;
                    }
                }
            }
            if (numFmtId != undefined) {
                var numf = numfmts[parseInt(numFmtId)];
                var cellFormat = new LuckyBase_1.LuckySheetCellFormat();
                cellFormat.fa = method_1.escapeCharacter(numf);
                // console.log(numf, numFmtId, this.v);
                cellFormat.t = t;
                cellValue.ct = cellFormat;
            }
            if (fillId != undefined) {
                var fillIdNum = parseInt(fillId);
                var fill = fills[fillIdNum];
                // console.log(cellValue.v);
                var bg = this.getBackgroundByFill(fill, clrScheme);
                if (bg != null) {
                    cellValue.bg = bg;
                }
            }
            if (fontId != undefined) {
                var fontIdNum = parseInt(fontId);
                var font = fonts[fontIdNum];
                if (font != null) {
                    var sz = font.getInnerElements("sz"); //font size
                    var colors = font.getInnerElements("color"); //font color
                    var family = font.getInnerElements("name"); //font family
                    var familyOverrides = font.getInnerElements("family"); //font family will be overrided by name
                    var charset = font.getInnerElements("charset"); //font charset
                    var bolds = font.getInnerElements("b"); //font bold
                    var italics = font.getInnerElements("i"); //font italic
                    var strikes = font.getInnerElements("strike"); //font italic
                    var underlines = font.getInnerElements("u"); //font italic
                    if (sz != null && sz.length > 0) {
                        var fs = sz[0].attributeList.val;
                        if (fs != null) {
                            cellValue.fs = parseInt(fs);
                        }
                    }
                    if (colors != null && colors.length > 0) {
                        var color = colors[0];
                        var fc = ReadXml_1.getColor(color, this.styles, "t");
                        if (fc != null) {
                            cellValue.fc = fc;
                        }
                    }
                    if (familyOverrides != null && familyOverrides.length > 0) {
                        var val = familyOverrides[0].attributeList.val;
                        if (val != null) {
                            familyFont = constant_1.fontFamilys[val];
                        }
                    }
                    if (family != null && family.length > 0) {
                        var val = family[0].attributeList.val;
                        if (val != null) {
                            cellValue.ff = val;
                        }
                    }
                    if (bolds != null && bolds.length > 0) {
                        var bold = bolds[0].attributeList.val;
                        if (bold == "0") {
                            cellValue.bl = 0;
                        }
                        else {
                            cellValue.bl = 1;
                        }
                    }
                    if (italics != null && italics.length > 0) {
                        var italic = italics[0].attributeList.val;
                        if (italic == "0") {
                            cellValue.it = 0;
                        }
                        else {
                            cellValue.it = 1;
                        }
                    }
                    if (strikes != null && strikes.length > 0) {
                        var strike = strikes[0].attributeList.val;
                        if (strike == "0") {
                            cellValue.cl = 0;
                        }
                        else {
                            cellValue.cl = 1;
                        }
                    }
                    if (underlines != null && underlines.length > 0) {
                        var underline = underlines[0].attributeList.val;
                        if (underline == "single") {
                            cellValue.un = 1;
                        }
                        else if (underline == "double") {
                            cellValue.un = 2;
                        }
                        else if (underline == "singleAccounting") {
                            cellValue.un = 3;
                        }
                        else if (underline == "doubleAccounting") {
                            cellValue.un = 4;
                        }
                        else {
                            cellValue.un = 0;
                        }
                    }
                }
            }
            // vt: number | undefined//Vertical alignment, 0 middle, 1 up, 2 down, alignment
            // ht: number | undefined//Horizontal alignment,0 center, 1 left, 2 right, alignment
            // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
            // tb: number | undefined //Text wrap,0 truncation, 1 overflow, 2 word wrap, alignment
            if (horizontal != undefined) { //Horizontal alignment
                if (horizontal == "center") {
                    cellValue.ht = 0;
                }
                else if (horizontal == "centerContinuous") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else if (horizontal == "left") {
                    cellValue.ht = 1;
                }
                else if (horizontal == "right") {
                    cellValue.ht = 2;
                }
                else if (horizontal == "distributed") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else if (horizontal == "fill") {
                    cellValue.ht = 1; //luckysheet unsupport
                }
                else if (horizontal == "general") {
                    cellValue.ht = 1; //luckysheet unsupport
                }
                else if (horizontal == "justify") {
                    cellValue.ht = 0; //luckysheet unsupport
                }
                else {
                    cellValue.ht = 1;
                }
            }
            if (vertical != undefined) { //Vertical alignment
                if (vertical == "bottom") {
                    cellValue.vt = 2;
                }
                else if (vertical == "center") {
                    cellValue.vt = 0;
                }
                else if (vertical == "distributed") {
                    cellValue.vt = 0; //luckysheet unsupport
                }
                else if (vertical == "justify") {
                    cellValue.vt = 0; //luckysheet unsupport
                }
                else if (vertical == "top") {
                    cellValue.vt = 1;
                }
                else {
                    cellValue.vt = 1;
                }
            }
            if (wrapText != undefined) {
                if (wrapText == "1") {
                    cellValue.tb = 2;
                }
                else {
                    cellValue.tb = 1;
                }
            }
            else {
                cellValue.tb = 1;
            }
            if (textRotation != undefined) {
                // tr: number | undefined //Text rotation,0: 0、1: 45 、2: -45、3 Vertical text、4: 90 、5: -90, alignment
                if (textRotation == "255") {
                    cellValue.tr = 3;
                }
                // else if(textRotation=="45"){
                //     cellValue.tr = 1;
                // }
                // else if(textRotation=="90"){
                //     cellValue.tr = 4;
                // }
                // else if(textRotation=="135"){
                //     cellValue.tr = 2;
                // }
                // else if(textRotation=="180"){
                //     cellValue.tr = 5;
                // }
                else {
                    cellValue.tr = 0;
                    cellValue.rt = parseInt(textRotation);
                }
            }
            if (shrinkToFit != undefined) { //luckysheet unsupport
            }
            if (indent != undefined) { //luckysheet unsupport
            }
            if (borderId != undefined) {
                var borderIdNum = parseInt(borderId);
                var border = borders[borderIdNum];
                // this._borderId = borderIdNum;
                var borderObject = new LuckyBase_1.LuckySheetborderInfoCellForImp();
                borderObject.rangeType = "cell";
                // borderObject.cells = [];
                var borderCellValue = new LuckyBase_1.LuckySheetborderInfoCellValue();
                borderCellValue.row_index = this.r;
                borderCellValue.col_index = this.c;
                var lefts = border.getInnerElements("left");
                var rights = border.getInnerElements("right");
                var tops = border.getInnerElements("top");
                var bottoms = border.getInnerElements("bottom");
                var diagonals = border.getInnerElements("diagonal");
                var starts = border.getInnerElements("start");
                var ends = border.getInnerElements("end");
                var left = this.getBorderInfo(lefts);
                var right = this.getBorderInfo(rights);
                var top_1 = this.getBorderInfo(tops);
                var bottom = this.getBorderInfo(bottoms);
                var diagonal = this.getBorderInfo(diagonals);
                var start = this.getBorderInfo(starts);
                var end = this.getBorderInfo(ends);
                var isAdd = false;
                if (start != null && start.color != null) {
                    borderCellValue.l = start;
                    isAdd = true;
                }
                if (end != null && end.color != null) {
                    borderCellValue.r = end;
                    isAdd = true;
                }
                if (left != null && left.color != null) {
                    borderCellValue.l = left;
                    isAdd = true;
                }
                if (right != null && right.color != null) {
                    borderCellValue.r = right;
                    isAdd = true;
                }
                if (top_1 != null && top_1.color != null) {
                    borderCellValue.t = top_1;
                    isAdd = true;
                }
                if (bottom != null && bottom.color != null) {
                    borderCellValue.b = bottom;
                    isAdd = true;
                }
                if (isAdd) {
                    borderObject.value = borderCellValue;
                    // this.config._borderInfo[borderId] = borderObject;
                    this._borderObject = borderObject;
                }
            }
        }
        else {
            cellValue.tb = 1;
        }
        if (v != null) {
            var value = v[0].value;
            if (/&#\d+;/.test(value)) {
                value = this.htmlDecode(value);
            }
            if (t == constant_1.ST_CellType["SharedString"]) {
                var siIndex = parseInt(v[0].value);
                var sharedSI = sharedStrings[siIndex];
                var rFlag = sharedSI.getInnerElements("r");
                if (rFlag == null) {
                    var tFlag = sharedSI.getInnerElements("t");
                    if (tFlag != null) {
                        var text_1 = "";
                        tFlag.forEach(function (t) {
                            text_1 += t.value;
                        });
                        text_1 = method_1.escapeCharacter(text_1);
                        //isContainMultiType(text) &&
                        if (familyFont == "Roman" && text_1.length > 0) {
                            var textArray = text_1.split("");
                            var preWordType = null, wordText = "", preWholef = null;
                            var wholef = "Times New Roman";
                            if (cellValue.ff != null) {
                                wholef = cellValue.ff;
                            }
                            var cellFormat = cellValue.ct;
                            if (cellFormat == null) {
                                cellFormat = new LuckyBase_1.LuckySheetCellFormat();
                            }
                            if (cellFormat.s == null) {
                                cellFormat.s = [];
                            }
                            for (var i = 0; i < textArray.length; i++) {
                                var w = textArray[i];
                                var type = null, ff = wholef;
                                if (method_1.isChinese(w)) {
                                    type = "c";
                                    ff = "宋体";
                                }
                                else if (method_1.isJapanese(w)) {
                                    type = "j";
                                    ff = "Yu Gothic";
                                }
                                else if (method_1.isKoera(w)) {
                                    type = "k";
                                    ff = "Malgun Gothic";
                                }
                                else {
                                    type = "e";
                                }
                                if ((type != preWordType && preWordType != null) || i == textArray.length - 1) {
                                    var InlineString = new LuckyBase_1.LuckyInlineString();
                                    InlineString.ff = preWholef;
                                    if (cellValue.fc != null) {
                                        InlineString.fc = cellValue.fc;
                                    }
                                    if (cellValue.fs != null) {
                                        InlineString.fs = cellValue.fs;
                                    }
                                    if (cellValue.cl != null) {
                                        InlineString.cl = cellValue.cl;
                                    }
                                    if (cellValue.un != null) {
                                        InlineString.un = cellValue.un;
                                    }
                                    if (cellValue.bl != null) {
                                        InlineString.bl = cellValue.bl;
                                    }
                                    if (cellValue.it != null) {
                                        InlineString.it = cellValue.it;
                                    }
                                    if (i == textArray.length - 1) {
                                        if (type == preWordType) {
                                            InlineString.ff = ff;
                                            InlineString.v = wordText + w;
                                        }
                                        else {
                                            InlineString.ff = preWholef;
                                            InlineString.v = wordText;
                                            cellFormat.s.push(InlineString);
                                            var InlineStringLast = new LuckyBase_1.LuckyInlineString();
                                            InlineStringLast.ff = ff;
                                            InlineStringLast.v = w;
                                            if (cellValue.fc != null) {
                                                InlineStringLast.fc = cellValue.fc;
                                            }
                                            if (cellValue.fs != null) {
                                                InlineStringLast.fs = cellValue.fs;
                                            }
                                            if (cellValue.cl != null) {
                                                InlineStringLast.cl = cellValue.cl;
                                            }
                                            if (cellValue.un != null) {
                                                InlineStringLast.un = cellValue.un;
                                            }
                                            if (cellValue.bl != null) {
                                                InlineStringLast.bl = cellValue.bl;
                                            }
                                            if (cellValue.it != null) {
                                                InlineStringLast.it = cellValue.it;
                                            }
                                            cellFormat.s.push(InlineStringLast);
                                            break;
                                        }
                                    }
                                    else {
                                        InlineString.v = wordText;
                                    }
                                    cellFormat.s.push(InlineString);
                                    wordText = w;
                                }
                                else {
                                    wordText += w;
                                }
                                preWordType = type;
                                preWholef = ff;
                            }
                            cellFormat.t = "inlineStr";
                            // cellFormat.s = [InlineString];
                            cellValue.ct = cellFormat;
                            // console.log(cellValue);
                        }
                        else {
                            text_1 = this.replaceSpecialWrap(text_1);
                            if (text_1.indexOf("\r\n") > -1 || text_1.indexOf("\n") > -1) {
                                var InlineString = new LuckyBase_1.LuckyInlineString();
                                InlineString.v = text_1;
                                var cellFormat = cellValue.ct;
                                if (cellFormat == null) {
                                    cellFormat = new LuckyBase_1.LuckySheetCellFormat();
                                }
                                if (cellValue.ff != null) {
                                    InlineString.ff = cellValue.ff;
                                }
                                if (cellValue.fc != null) {
                                    InlineString.fc = cellValue.fc;
                                }
                                if (cellValue.fs != null) {
                                    InlineString.fs = cellValue.fs;
                                }
                                if (cellValue.cl != null) {
                                    InlineString.cl = cellValue.cl;
                                }
                                if (cellValue.un != null) {
                                    InlineString.un = cellValue.un;
                                }
                                if (cellValue.bl != null) {
                                    InlineString.bl = cellValue.bl;
                                }
                                if (cellValue.it != null) {
                                    InlineString.it = cellValue.it;
                                }
                                cellFormat.t = "inlineStr";
                                cellFormat.s = [InlineString];
                                cellValue.ct = cellFormat;
                            }
                            else {
                                cellValue.v = text_1;
                                quotePrefix = "1";
                            }
                        }
                    }
                }
                else {
                    var styles_1 = [];
                    rFlag.forEach(function (r) {
                        var tFlag = r.getInnerElements("t");
                        var rPr = r.getInnerElements("rPr");
                        var InlineString = new LuckyBase_1.LuckyInlineString();
                        if (tFlag != null && tFlag.length > 0) {
                            var text = tFlag[0].value;
                            text = _this.replaceSpecialWrap(text);
                            text = method_1.escapeCharacter(text);
                            InlineString.v = text;
                        }
                        if (rPr != null && rPr.length > 0) {
                            var frpr = rPr[0];
                            var sz = ReadXml_1.getlineStringAttr(frpr, "sz"), rFont = ReadXml_1.getlineStringAttr(frpr, "rFont"), family = ReadXml_1.getlineStringAttr(frpr, "family"), charset = ReadXml_1.getlineStringAttr(frpr, "charset"), scheme = ReadXml_1.getlineStringAttr(frpr, "scheme"), b = ReadXml_1.getlineStringAttr(frpr, "b"), i = ReadXml_1.getlineStringAttr(frpr, "i"), u = ReadXml_1.getlineStringAttr(frpr, "u"), strike = ReadXml_1.getlineStringAttr(frpr, "strike"), vertAlign = ReadXml_1.getlineStringAttr(frpr, "vertAlign"), color = void 0;
                            var cEle = frpr.getInnerElements("color");
                            if (cEle != null && cEle.length > 0) {
                                color = ReadXml_1.getColor(cEle[0], _this.styles, "t");
                            }
                            var ff = void 0;
                            // if(family!=null){
                            //     ff = fontFamilys[family];
                            // }
                            if (rFont != null) {
                                ff = rFont;
                            }
                            if (ff != null) {
                                InlineString.ff = ff;
                            }
                            else if (cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }
                            if (color != null) {
                                InlineString.fc = color;
                            }
                            else if (cellValue.fc != null) {
                                InlineString.fc = cellValue.fc;
                            }
                            if (sz != null) {
                                InlineString.fs = parseInt(sz);
                            }
                            else if (cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }
                            if (strike != null) {
                                InlineString.cl = parseInt(strike);
                            }
                            else if (cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }
                            if (u != null) {
                                InlineString.un = parseInt(u);
                            }
                            else if (cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }
                            if (b != null) {
                                InlineString.bl = parseInt(b);
                            }
                            else if (cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }
                            if (i != null) {
                                InlineString.it = parseInt(i);
                            }
                            else if (cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }
                            if (vertAlign != null) {
                                InlineString.va = parseInt(vertAlign);
                            }
                            // ff:string | undefined //font family
                            // fc:string | undefined//font color
                            // fs:number | undefined//font size
                            // cl:number | undefined//strike
                            // un:number | undefined//underline
                            // bl:number | undefined//blod
                            // it:number | undefined//italic
                            // v:string | undefined
                        }
                        else {
                            if (InlineString.ff == null && cellValue.ff != null) {
                                InlineString.ff = cellValue.ff;
                            }
                            if (InlineString.fc == null && cellValue.fc != null) {
                                InlineString.fc = cellValue.fc;
                            }
                            if (InlineString.fs == null && cellValue.fs != null) {
                                InlineString.fs = cellValue.fs;
                            }
                            if (InlineString.cl == null && cellValue.cl != null) {
                                InlineString.cl = cellValue.cl;
                            }
                            if (InlineString.un == null && cellValue.un != null) {
                                InlineString.un = cellValue.un;
                            }
                            if (InlineString.bl == null && cellValue.bl != null) {
                                InlineString.bl = cellValue.bl;
                            }
                            if (InlineString.it == null && cellValue.it != null) {
                                InlineString.it = cellValue.it;
                            }
                        }
                        styles_1.push(InlineString);
                    });
                    var cellFormat = cellValue.ct;
                    if (cellFormat == null) {
                        cellFormat = new LuckyBase_1.LuckySheetCellFormat();
                    }
                    cellFormat.t = "inlineStr";
                    cellFormat.s = styles_1;
                    cellValue.ct = cellFormat;
                }
            }
            // else if(t==ST_CellType["InlineString"] && v!=null){
            // }
            else {
                value = method_1.escapeCharacter(value);
                cellValue.v = value;
            }
        }
        if (quotePrefix != null) {
            cellValue.qp = parseInt(quotePrefix);
        }
        return cellValue;
    };
    LuckySheetCelldata.prototype.replaceSpecialWrap = function (text) {
        text = text.replace(/_x000D_/g, "").replace(/&#13;&#10;/g, "\r\n").replace(/&#13;/g, "\r").replace(/&#10;/g, "\n");
        return text;
    };
    LuckySheetCelldata.prototype.getBackgroundByFill = function (fill, clrScheme) {
        var patternFills = fill.getInnerElements("patternFill");
        if (patternFills != null) {
            var patternFill = patternFills[0];
            var fgColors = patternFill.getInnerElements("fgColor");
            var bgColors = patternFill.getInnerElements("bgColor");
            var fg = void 0, bg = void 0;
            if (fgColors != null) {
                var fgColor = fgColors[0];
                fg = ReadXml_1.getColor(fgColor, this.styles);
            }
            if (bgColors != null) {
                var bgColor = bgColors[0];
                bg = ReadXml_1.getColor(bgColor, this.styles);
            }
            // console.log(fgColors,bgColors,clrScheme);
            if (fg != null) {
                return fg;
            }
            else if (bg != null) {
                return bg;
            }
        }
        else {
            var gradientfills = fill.getInnerElements("gradientFill");
            if (gradientfills != null) {
                //graient color fill handler
                return null;
            }
        }
    };
    LuckySheetCelldata.prototype.getBorderInfo = function (borders) {
        if (borders == null) {
            return null;
        }
        var border = borders[0], attrList = border.attributeList;
        var clrScheme = this.styles["clrScheme"];
        var style = attrList.style;
        if (style == null || style == "none") {
            return null;
        }
        var colors = border.getInnerElements("color");
        var colorRet = "#000000";
        if (colors != null) {
            var color = colors[0];
            colorRet = ReadXml_1.getColor(color, this.styles, "b");
            if (colorRet == null) {
                colorRet = "#000000";
            }
        }
        var ret = new LuckyBase_1.LuckySheetborderInfoCellValueStyle();
        ret.style = constant_1.borderTypes[style];
        ret.color = colorRet;
        return ret;
    };
    LuckySheetCelldata.prototype.htmlDecode = function (str) {
        return str.replace(/&#(x)?([^&]{1,5});?/g, function ($, $1, $2) {
            return String.fromCharCode(parseInt($2, $1 ? 16 : 10));
        });
    };
    ;
    return LuckySheetCelldata;
}(LuckyBase_1.LuckySheetCelldataBase));
exports.LuckySheetCelldata = LuckySheetCelldata;
