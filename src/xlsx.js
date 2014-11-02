var fs     = require('fs'),
    path   = require('path'),
    async  = require('async'),
    AdmZip = require('adm-zip'),

    blobs = require('./blobs'),

    numberRegex = /^[1-9\.][\d\.]+$/;

exports = module.exports = XlsxWriter;

function XlsxWriter() {
    this.strings = [];
    this.stringMap = {};
    this.stringIndex = 0;
    this.currentRow = 0;

    this.haveHeader = false;

    this.sheetBuffers = [];

    this.cellMap = [];
    this.cellLabelMap = {};
}

XlsxWriter.prototype.pack = function(filename, cb) {
    var dimensions = this.dimensions(this.currentRow, this.cellMap.length),
        zip = new AdmZip(),
        self = this;

    return async.series([
            function(cb) {
                var string, stringTable, _i, _len, _ref;
                stringTable = '';
                _ref = self.strings;
                for (_i = 0, _len = _ref.length; _i < _len; _i++) {
                    string = _ref[_i];
                    stringTable += blobs.string(self.escapeXml(string));
                }
                zip.addFile('xl/sharedStrings.xml', new Buffer(blobs.stringsHeader(self.strings.length) + stringTable + blobs.stringsFooter));

                cb();
            }, function(cb) {
                zip.addFile('[Content_Types].xml', new Buffer(blobs.contentTypes));

                cb();
            }, function(cb) {
                zip.addFile('_rels/.rels', new Buffer(blobs.rels));

                cb();
            }, function(cb) {
                zip.addFile('xl/workbook.xml', new Buffer(blobs.workbook));

                cb();
            }, function(cb) {
                zip.addFile('xl/styles.xml', new Buffer(blobs.styles));

                cb();
            }, function(cb) {
                zip.addFile('xl/_rels/workbook.xml.rels', new Buffer(blobs.workbookRels));

                cb();
            }, function(cb) {
                var buffers = Array.prototype.concat([new Buffer(blobs.sheetHeader(dimensions))], self.sheetBuffers, [new Buffer(blobs.sheetFooter)]);
                zip.addFile('xl/worksheets/sheet1.xml', Buffer.concat(buffers));

                cb();
            }, function(cb) {
                zip.writeZip(filename);

                cb();
            }
    ], cb);
};

XlsxWriter.prototype.addRow = function addRow(obj) {
    if (!this.haveHeader) {
        this._startRow();
        var col = 1;

        for (var key in obj) {
            this._addCell(key, col);
            this.cellMap.push(key);
            col += 1;
        }

        this._endRow();
        this.haveHeader = true;
    }

    this._startRow();
    for (var col = 0, len = this.cellMap.length; col < len; col += 1) {
        this._addCell(obj[this.cellMap[col]] || "", col + 1);
    }
    return this._endRow();
};

XlsxWriter.prototype._addCell = function(value, col) {
    var cell, index, row;

    if (value == null) {
        value = '';
    }

    row = this.currentRow;
    cell = this.cell(row, col);

    if (numberRegex.test(value)) {
        return this.rowBuffer += blobs.numberCell(value, cell);
    } else {
        index = this._lookupString(value);
        return this.rowBuffer += blobs.cell(index, cell);
    }
};

XlsxWriter.prototype.dimensions = function dimensions(rows, columns) {
    return 'A1:' + this.cell(rows, columns);
}

XlsxWriter.prototype.cell = function cell(row, col) {
    var colIndex = '';

    if (this.cellLabelMap[col]) {
        colIndex = this.cellLabelMap[col];
    } else {
        if (col == 0) {
            // Provide a fallback for empty spreadsheets
            row = 1
            col = 1
        }

        input = (+col - 1).toString(26);
        while (input.length) {
            a = input.charCodeAt(input.length - 1);
            colIndex = String.fromCharCode(a + (a >= 48 && a <= 57 ? 17 : -22)) + colIndex;
            input = input.length > 1 ? (parseInt(input.substr(0, input.length - 1), 26) - 1).toString(26) : "";
        }

        this.cellLabelMap[col] = colIndex;
    }

    return colIndex + row;
}

XlsxWriter.prototype._startRow = function() {
    this.rowBuffer = blobs.startRow(this.currentRow);
    return this.currentRow += 1;
};

XlsxWriter.prototype._lookupString = function(value) {
    if (!this.stringMap[value]) {
        this.stringMap[value] = this.stringIndex;
        this.strings.push(value);
        this.stringIndex += 1;
    }

    return this.stringMap[value];
};

XlsxWriter.prototype._endRow = function() {
    return this.sheetBuffers.push(new Buffer(this.rowBuffer + blobs.endRow));
};

XlsxWriter.prototype.escapeXml = function(str) {
    if (str == null) {
        str = '';
    }

    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
};
