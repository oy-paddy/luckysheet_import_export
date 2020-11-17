"use strict";
exports.__esModule = true;
exports.HandleZip = void 0;
var jszip_1 = require("jszip");
var method_1 = require("./common/method");
var HandleZip = /** @class */ (function () {
    function HandleZip(file) {
        if (file instanceof File) {
            this.uploadFile = file;
        }
    }
    HandleZip.prototype.unzipFile = function (successFunc, errorFunc) {
        // var new_zip:JSZip = new JSZip();
        jszip_1["default"].loadAsync(this.uploadFile) // 1) read the Blob
            .then(function (zip) {
            var fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
            zip.forEach(function (relativePath, zipEntry) {
                var fileName = zipEntry.name;
                var fileNameArr = fileName.split(".");
                var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                var fileType = "string";
                if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1 }) {
                    fileType = "base64";
                }
                else if (suffix == "emf") {
                    fileType = "arraybuffer";
                }
                zipEntry.async(fileType).then(function (data) {
                    if (fileType == "base64") {
                        data = "data:image/" + suffix + ";base64," + data;
                    }
                    fileList[zipEntry.name] = data;
                    // console.log(lastIndex, index);
                    if (lastIndex == index + 1) {
                        successFunc(fileList);
                    }
                    index++;
                });
            });
        }, function (e) {
            errorFunc(e);
        });
    };
    HandleZip.prototype.unzipFileByUrl = function (url, successFunc, errorFunc) {
        var new_zip = new jszip_1["default"]();
        method_1.getBinaryContent(url, function (err, data) {
            if (err) {
                throw err; // or handle err
            }
            jszip_1["default"].loadAsync(data).then(function (zip) {
                var fileList = {}, lastIndex = Object.keys(zip.files).length, index = 0;
                zip.forEach(function (relativePath, zipEntry) {
                    var fileName = zipEntry.name;
                    var fileNameArr = fileName.split(".");
                    var suffix = fileNameArr[fileNameArr.length - 1].toLowerCase();
                    var fileType = "string";
                    if (suffix in { "png": 1, "jpeg": 1, "jpg": 1, "gif": 1, "bmp": 1, "tif": 1, "webp": 1 }) {
                        fileType = "base64";
                    }
                    else if (suffix == "emf") {
                        fileType = "arraybuffer";
                    }
                    zipEntry.async(fileType).then(function (data) {
                        if (fileType == "base64") {
                            data = "data:image/" + suffix + ";base64," + data;
                        }
                        fileList[zipEntry.name] = data;
                        // console.log(lastIndex, index);
                        if (lastIndex == index + 1) {
                            successFunc(fileList);
                        }
                        index++;
                    });
                });
            }, function (e) {
                errorFunc(e);
            });
        });
    };
    HandleZip.prototype.newZipFile = function () {
        var zip = new jszip_1["default"]();
        this.workBook = zip;
    };
    //title:"nested/hello.txt", content:"Hello Worldasdfasfasdfasfasfasfasfasdfas"
    HandleZip.prototype.addToZipFile = function (title, content) {
        if (this.workBook == null) {
            var zip = new jszip_1["default"]();
            this.workBook = zip;
        }
        this.workBook.file(title, content);
    };
    return HandleZip;
}());
exports.HandleZip = HandleZip;
