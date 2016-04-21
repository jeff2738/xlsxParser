var xls = require('node-xlsx');

var Service = function (xlsFile, titleIdx) {
  this.xlsObject = xls.parse(xlsFile);
  this.titleIdx  = titleIdx;
}

Service.prototype.getByTitle = function (columnTitle, row) {
  var sheet    = this.xlsObject[0].data;
  var titleIdx = this.titleIdx;
  var colIdx   = sheet[titleIdx].indexOf(columnTitle);
  var row = sheet[row];
  if(!row){
    return "";
  }
  var content = row[colIdx];
  if(typeof(content) == "String" ){
    return content.trim();
  }else{
    return content;
  }

}
module.exports               = Service;