offin = require("../index.js")

console.log offin

data =
  firstName: "Joe"
  lastName: "Dirt"


  
offin.docx 
  template: "./templates/template.docx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.docx"
  data: data
  
offin.pptx
  template: "./templates/template.pptx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.pptx"
  data: data

offin.xlsx  
  template: "./templates/template.xlsx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.xlsx"
  data: data