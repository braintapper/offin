// module.exports = require("./lib/injector_wrapper").injector
var Docxtemplater, Injector, PizZip, Sugar, fs;

Sugar = require("sugar-and-spice");

Sugar.extend();

fs = require('fs');

PizZip = require('pizzip');

Docxtemplater = require('docxtemplater');

Injector = {
  docx: function(options) {
    var buf, content, doc, zip;
    console.log(options);
    content = fs.readFileSync(options.template, 'binary');
    zip = new PizZip(content);
    doc = new Docxtemplater(zip); //.setOptions {}
    doc.setData(options.data);
    doc.render();
    buf = doc.getZip().generate({
      type: 'nodebuffer'
    });
    return fs.writeFileSync(options.output, buf);
  },
  pptx: function(options) {
    var buf, content, doc, zip;
    console.log(options);
    content = fs.readFileSync(options.template, 'binary');
    zip = new PizZip(content);
    doc = new Docxtemplater(zip); //.setOptions {}
    doc.setData(options.data);
    doc.render();
    buf = doc.getZip().generate({
      type: 'nodebuffer'
    });
    return fs.writeFileSync(options.output, buf);
  },
  xlsx: function(options) {
    var XlsxTemplate, content, workbook;
    // this require is here because docxtemplater can't coexist with sugar
    console.log(options);
    XlsxTemplate = require('./xlsx.js');
    content = fs.readFileSync(options.template, 'binary');
    workbook = new XlsxTemplate(content);
    // iterate into all sheets
    workbook.sheets.forEach(function(sheet) {
      return workbook.substitute(sheet.id, options.data);
    });
    return workbook.writeFile(options.output);
  },
  inject: function(options) {}
};

module.exports = Injector;
