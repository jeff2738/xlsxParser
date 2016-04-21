var Service   = require('./index');

var service = new Service("excel.xlsx",0);
var doc1=service.getByTitle("BBB",3);
console.log(doc1);
var doc2=service.getByTitle("CCC",1);
console.log(doc2);