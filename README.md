node-excel-api [![Build Status](https://secure.travis-ci.org/layerssss/node-excel-api.png)](http://travis-ci.org/layerssss/node-excel-api)
======================

**node-excel-api** is a simple wrapper arround the mature [JExcelApi](http://jexcelapi.sourceforge.net/) project. So most APIs of JExcelApi are left unchanged. You can view documentation in [JExcelApi's javadoc](http://jexcelapi.sourceforge.net/resources/javadocs/2_6_10/docs/index.html)

Some methods are modified a little bit, see [lib/excel.iced](lib/excel.iced)

Examples
-----------------------

### Parsing

```
Workbook = require('excel-api').Workbook;
Workbook.getWorkbook('./input.xls', function(err, workbook){

  workbook.getSheetNames(function(err, names){
    console.log('sheets: ' + names);

    workbook.getSheet('Sheet1', function(err, sheet){

      sheet.getRows(function(err, rows){
        console.log('rows: ' + rows);

        for(var i = 0; i < rows; i++){
          sheet.getRow(function(err, cells){
            for(var j = 0; j < cells.length; j++){
              cells[j].getContents(function(err, content){
                console.log(i + 'x' + j + ': ' + content);
              });
            }
          });
        }
      });
    });
  });
});
```