offin = require("../index.js")

console.log offin

data =
  firstName: "Joe"
  lastName: "Dirt"

options =
  template: "./templates/template.docx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.docx"
  data: data
  
#offin.docx options 

options =
  template: "./templates/template.pptx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.pptx"
  data: data
  
#offin.pptx options 


options =
  template: "./templates/template.xlsx"
  output: "./output/output-#{Date.create().format("{HH}{mm}{ss}")}.xlsx"
  data: data
offin.xlsx options 