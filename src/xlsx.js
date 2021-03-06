var fs     = require('fs'),
    path   = require('path'),
    AdmZip = require('adm-zip'),

    blobs  = require('./blobs');

const NUMBER_REGEX = /^\-?[1-9\.][\d\.]*$/;

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

XlsxWriter.write = function(filename, data, callback) {
    var writer = new XlsxWriter();

    data.forEach(function(row) {
        writer.addRow(row);
    });

    writer.pack(filename, callback);
};

XlsxWriter.prototype.pack = function(filename, callback) {
    var dimensions = this.dimensions(this.currentRow, this.cellMap.length),
        zip = new AdmZip(),
        self = this;

    // xl/sharedstrings.xml
    var stringTable = '';
    this.strings.forEach(function(string) {
        stringTable += blobs.string(self.escapeXml(string));
    });
    zip.addFile('xl/sharedStrings.xml', new Buffer(blobs.stringsHeader(this.strings.length) + stringTable + blobs.stringsFooter));

    // [Content_types].xml
    zip.addFile('[Content_Types].xml', new Buffer(blobs.contentTypes));

    // _rels/.rels
    zip.addFile('_rels/.rels', new Buffer(blobs.rels));

    // xl/workbook.xml
    zip.addFile('xl/workbook.xml', new Buffer(blobs.workbook));

    // xl/styles.xml
    zip.addFile('xl/styles.xml', new Buffer(blobs.styles));

    // xl/_rels/workbook.xml.rels
    zip.addFile('xl/_rels/workbook.xml.rels', new Buffer(blobs.workbookRels));

    // xl/worksheets/sheet1.xml
    var buffers = Array.prototype.concat([new Buffer(blobs.sheetHeader(dimensions))], this.sheetBuffers, [new Buffer(blobs.sheetFooter)]);
    zip.addFile('xl/worksheets/sheet1.xml', Buffer.concat(buffers));

    // write the zip to the filesystem
    zip.writeZip(filename);

    callback();
};

XlsxWriter.prototype.setHeaders = function setHeaders(headers) {
    if (this.haveHeader) {
        throw new Error('Headers have already been set');
    }

    var col  = 1,
        self = this;

    this._startRow();

    headers.forEach(function(header) {
        self._addCell(header, col);
        self.cellMap.push(header);
        col += 1;
    });

    this._endRow();
    this.haveHeader = true;
};

XlsxWriter.prototype.addRow = function addRow(obj) {
    if (!this.haveHeader) {
        var headers = [];
        for (var key in obj) {
            headers.push(key);
        }

        this.setHeaders(headers);
    }

    var col = 0;

    this._startRow();
    for (var key in obj) {
        this._addCell(obj[key] || "", col + 1);

        col += 1;
    }
    this._endRow();
};

XlsxWriter.prototype._addCell = function(value, col) {
    var cell, index, row;

    if (value == null) {
        value = '';
    }

    row = this.currentRow;
    cell = this.cell(row, col);

    if (NUMBER_REGEX.test(value)) {
        return this.rowBuffer += blobs.numberCell(value, cell);
    } else {
        index = this._lookupString(value);
        return this.rowBuffer += blobs.cell(index, cell);
    }
};

XlsxWriter.prototype.dimensions = function dimensions(rows, columns) {
    return 'A1:' + this.cell(rows, columns);
};

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
};

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
