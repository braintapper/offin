# Offin

A simple library for injecting JSON data into Microsoft Office XML document templates - `docx`, `xlsx`, `pptx`

This uses `docxtemplater` and `xslx-template` to accomplish this. For consistency, all file types use mustache style tokens:

```
{firstName} {lastName}
```

To see how templating works:

* DOCX and PPTX: https://docxtemplater.com/demo/#simple
* XLSX: 



# Installation

`npm install offin [--save]`






# Sample Code (coffeescript)

## Excel

```coffeescript

offin = require("offin")

data =
  firstName: "Joe"
  lastName: "Dirt"

options =
  template: "./template.xlsx"
  output: "./output.xlsx"
  data: data
  
offin.xlsx options 

```

## Word

```coffeescript

offin = require("offin")

data =
  firstName: "Joe"
  lastName: "Dirt"

options =
  template: "./template.docx"
  output: "./output.docx"
  data: data
  
offin.docx options 

```

## Powerpoint

```coffeescript

offin = require("offin")

data =
  firstName: "Joe"
  lastName: "Dirt"

options =
  template: "./template.pptx"
  output: "./output.pptx"
  data: data
  
offin.pptx options 

```




# Changelog



## 0.0.1

* Initial release



# Open Source Attributions

* DOCX and PPTX Output: https://github.com/open-xml-templating/docxtemplater
* XLSX Output: https://github.com/optilude/xlsx-template