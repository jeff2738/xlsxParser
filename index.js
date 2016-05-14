var xls = require('node-xlsx');

var Service = function (xlsFile, titleIdx) {
  this.xlsObject = xls.parse(xlsFile);
  this.titleIdx  = titleIdx;
}

Service.prototype.getByTitle = function (columnTitle, row) {
  var sheet    = this.xlsObject[0].data;
  var titleIdx = this.titleIdx;
  var colIdx   = sheet[titleIdx].indexOf(columnTitle);
  var row      = sheet[row];
  if (!row) {
    return null;
  }
  var content = row[colIdx];
  if (typeof(content) == "String") {
    return content.trim();
  } else if (content == "-") {
    return "";
  } else {
    return content;
  }
}

Service.prototype.getLength = function () {
  var sheet = this.xlsObject[0].data;
  for (i = 0; i < sheet.length; i++) {
    var row = sheet[i];
    if (row == undefined || row.length == 0) {
      return i;
    }
  }
  return sheet.length;
}

module.exports = Service;