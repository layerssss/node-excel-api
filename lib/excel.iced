java = require 'java'
path = require 'path'


java.classpath.push path.join  __dirname, '..', 'deps', 'jxl.jar'

jFile = java.import 'java.io.File'

exports.Workbook = Workbook = java.import 'jxl.Workbook'
jGetWorkbook = Workbook.getWorkbook
Workbook.getWorkbook = (path, args..., cb)->
  jGetWorkbook.call @, (new jFile path), args..., cb

jCreateWorkbook = Workbook.createWorkbook
Workbook.createWorkbook = (path, args..., cb)->
  jCreateWorkbook.call @, (new jFile path), args..., cb



exports.Sheet = Sheet = java.import 'jxl.Sheet'

exports.WorkbookSettings = WorkbookSettings = java.import 'jxl.WorkbookSettings'

exports.SheetSettings = SheetSettings = java.import 'jxl.SheetSettings'

