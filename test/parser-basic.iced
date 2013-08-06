assert = require 'assert'
path = require 'path'

{
  Workbook
} = require '../lib/excel'

describe 'parser-basic', ->
  before (cb)->
    await Workbook.getWorkbook (path.join __dirname, 'fixtures', 'parse-basic.xls'), defer e, @workbook
    return cb e if e
    cb()


  describe 'getSheetNames', ->
    it 'should contains the right names', (cb)->
      await @workbook.getSheetNames defer e, names
      return cb e if e
      assert.equal names[0], 'Sheet1'
      assert.equal names[1], 'Sheet2'
      assert.equal names[2], 'Sheet3'
      cb()

  describe 'getNumberOfSheets', ->
    it 'should ', (cb)->
      await @workbook.getNumberOfSheets defer e, num
      return cb e if e
      assert.equal num, 3
      cb()

  describe 'getSheet(0)', ->
    it '', (cb)->
      await @workbook.getSheet 0, defer e, sheet
      return cb e if e
      await sheet.getName defer e, name
      return cb e if e
      assert.equal name, 'Sheet1'
      cb()

  describe 'Sheet1', ->
    before (cb)->
      await @workbook.getSheet 'Sheet1', defer e, @Sheet1
      cb e

    describe 'getColumns', ->
      it '', (cb)->
        await @Sheet1.getColumns defer e, num
        return cb e if e
        assert.equal num, 3
        cb()
    describe 'getColumn(0)', ->
      it '', (cb)->
        await @Sheet1.getColumn 0, defer e, cells
        return cb e if e
        assert.equal cells.length, 1
        await cells[0].getContents defer e, content
        return cb e if e
        assert.equal content, 'a'
        cb()

    describe 'getRows', ->
      it '', (cb)->
        await @Sheet1.getRows defer e, num
        return cb e if e
        assert.equal num, 1
        cb()

    describe 'getRow(0)', ->
      it '', (cb)->
        await @Sheet1.getRow 0, defer e, cells
        return cb e if e
        assert.equal cells.length, 3
        await cells[0].getContents defer e, content
        return cb e if e
        assert.equal content, 'a'
        await cells[1].getContents defer e, content
        return cb e if e
        assert.equal content, 'b'
        await cells[2].getContents defer e, content
        return cb e if e
        assert.equal content, 'c'
        cb()
    
    

  after (cb)->
    await @workbook.close defer e
    return cb e if e
    cb()

