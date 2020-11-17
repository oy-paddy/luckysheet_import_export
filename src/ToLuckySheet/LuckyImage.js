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
exports.ImageList = void 0;
var LuckyBase_1 = require("./LuckyBase");
var emf_1 = require("../common/emf");
var ImageList = /** @class */ (function () {
    function ImageList(files) {
        if (files == null) {
            return;
        }
        this.images = {};
        for (var fileKey in files) {
            // let reg = new RegExp("xl/media/image1.png", "g");
            if (fileKey.indexOf("xl/media/") > -1) {
                var fileNameArr = fileKey.split(".");
                var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1, "emf": 1 }) {
                    if (suffix == "emf") {
                        var pNum = 0; // number of the page, that you want to render
                        var scale = 1; // the scale of the document
                        var wrt = new emf_1.ToContext2D(pNum, scale);
                        var inp, out, stt;
                        emf_1.FromEMF.K = [];
                        inp = emf_1.FromEMF.C;
                        out = emf_1.FromEMF.K;
                        stt = 4;
                        for (var p in inp)
                            out[inp[p]] = p.slice(stt);
                        emf_1.FromEMF.Parse(files[fileKey], wrt);
                        this.images[fileKey] = wrt.canvas.toDataURL("image/png");
                    }
                    else {
                        this.images[fileKey] = files[fileKey];
                    }
                }
            }
        }
    }
    ImageList.prototype.getImageByName = function (pathName) {
        if (pathName in this.images) {
            var base64 = this.images[pathName];
            return new Image(pathName, base64);
        }
        return null;
    };
    return ImageList;
}());
exports.ImageList = ImageList;
var Image = /** @class */ (function (_super) {
    __extends(Image, _super);
    function Image(pathName, base64) {
        var _this = _super.call(this) || this;
        _this.src = base64;
        return _this;
    }
    Image.prototype.setDefault = function () {
    };
    return Image;
}(LuckyBase_1.LuckyImageBase));
