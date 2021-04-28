




# module.exports = require("./lib/injector_wrapper").injector


Sugar = require "sugar-and-spice"
Sugar.extend()
fs = require('fs')
PizZip = require('pizzip')

Docxtemplater = require('docxtemplater')




Injector = {

  docx: (options)->
    console.log options
    content = fs.readFileSync(options.template, 'binary')
    zip = new PizZip(content)
    doc = new Docxtemplater(zip) #.setOptions {}
    doc.setData options.data
    doc.render()
    buf = doc.getZip().generate(type: 'nodebuffer')
    fs.writeFileSync options.output, buf

  pptx: (options)->
    console.log options
    content = fs.readFileSync(options.template, 'binary')
    zip = new PizZip(content)
    doc = new Docxtemplater(zip) #.setOptions {}
    doc.setData options.data
    doc.render()
    buf = doc.getZip().generate(type: 'nodebuffer')
    fs.writeFileSync options.output, buf

  xlsx: (options)->
    # this require is here because docxtemplater can't coexist with sugar
    console.log options    
    XlsxTemplate = require('./xlsx.js')
    content = fs.readFileSync(options.template, 'binary')
    workbook = new XlsxTemplate(content)
    # iterate into all sheets
    workbook.sheets.forEach (sheet)->
      workbook.substitute sheet.id, options.data
    workbook.writeFile(options.output)

  inject: (options)->

} 





module.exports = Injector
