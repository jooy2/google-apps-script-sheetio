class SheetMan {
  constructor () {
    this.Utils = new Utils();
    this.SheetApp = SpreadsheetApp;
    this.Sheets = Sheets;
    this.activeSheet = this.SheetApp.getActiveSpreadsheet();
    this.originSheet = this.activeSheet;
    this.sheet = null;
  }

  createFile (title, nameForFirstSheet) {
    this.sheet = this.SheetApp.create(title);

    if (arguments.length === 2) {
      this.sheet.getActiveSheet().setName(nameForFirstSheet);
    }

    this.sheet.createdSheetId = this.Utils.getFileIdByUrl(this.sheet.getUrl());

    return this;
  }

  createFileUsingApi (title) {
    const sheet = this.Sheets.newSpreadsheet();
    sheet.properties = this.Sheets.newSpreadsheetProperties();
    sheet.properties.title = title;

    const newFile = this.Sheets.SpreadSheets.create(sheet);
    this.sheet.createdSheetId = newFile.spreadsheetId;

    return this;
  }

  getFileId () {
    return this.sheet.createdSheetId
      ? this.sheet.createdSheetId
      : this.SheetApp.getActiveSpreadsheet().getId();
  }

  getSheetId () {
    return this.sheet.getSheetId();
  }

  getSheetName () {
    return this.sheet.getName();
  }

  targetTo (sheetId) {
    if (sheetId === 'self') {
      this.isExternalSheet = false;
      this.sheet = this.originSheet;
    } else {
      this.isExternalSheet = true;
      this.sheet = this.SheetApp.openById(sheetId);
    }

    this.activeSheet = this.sheet;

    return this;
  }

  targetSelf () {
    this.isExternalSheet = false;
    this.sheet = this.originSheet;

    this.activeSheet = this.sheet;

    return this;
  }

  active (sheetName) {
    this.sheet = this.activeSheet.getSheetByName(sheetName);
    if (this.sheet) this.sheet.name = sheetName;

    if (sheetName.length < 1 || !this.sheet) {
      throw `'${sheetName}' 시트가 존재하지 않습니다. 생성하려면 create()를 사용합니다.`;
    }

    return this;
  }

  isExist (sheetName) {
    if (this.isExternalSheet) {
      return this.sheet ? this.sheet.getSheetByName(sheetName) : null;
    }
    return this.activeSheet.getSheetByName(sheetName);
  }

  create (sheetName) {
    try {
      this.sheet = this.activeSheet.getSheetByName(sheetName);

      if (!this.sheet) {
        this.sheet = this.activeSheet.insertSheet(sheetName).activate();
      }
    } catch (e) {
      throw `'${sheetName}' 시트 생성에 실패했습니다. 이미 존재하거나 알 수 없는 에러입니다.`;
    }
    return this;
  }

  rename (renameTo) {
    if (!renameTo || renameTo.length < 1) {
      throw `'${sheetName}' 시트를 '${renameTo}'로 변경할 수 없습니다.`;
    }

    this.sheet.setName(renameTo);

    return this;
  }

  destroy () {
    try {
      this.SheetApp.getActive().deleteSheet(this.sheet);
    } catch (e) {
      throw `'${this.sheet.name}' 시트 삭제에 실패했습니다.`;
    }

    return this;
  }

  destroyByName (sheetName) {
    return this.active(sheetName).destroy();
  }

  flush () {
    this.SheetApp.flush();

    return this;
  }

  getId () {
    return this.SheetApp.getActiveSpreadsheet().getId();
  }

  getSheetCount () {
    return this.activeSheet.getSheets().length;
  }

  getRange (startRow, startColumn, rows, columns) {
    if (arguments.length === 1) {
      this.sheet.activeRange = this.sheet.getRange(startRow);
    } else if (arguments.length === 2) {
      this.sheet.activeRange = this.sheet.getRange(startRow, startColumn);
    } else {
      this.sheet.activeRange = this.sheet.getRange(startRow, startColumn, rows, columns);
    }

    return this;
  }

  getDataRange () {
    this.sheet.activeRange = this.sheet.getDataRange();

    return this;
  }

  getLastRow () {
    const lastRow = this.sheet.getLastRow();

    return lastRow === 0 ? 1 : lastRow;
  }

  getLastColumn () {
    const lastColumn = this.sheet.getLastColumn();

    return lastColumn === 0 ? 1 : lastColumn;
  }

  getMaxRows () {
    return this.sheet.getMaxRows();
  }

  getMaxColumns () {
    return this.sheet.getMaxColumns();
  }

  getFrozenRows () {
    return this.sheet.getFrozenRows();
  }

  getFrozenColumns () {
    return this.sheet.getFrozenColumns();
  }

  getValue () {
    return this.sheet.activeRange.getValue();
  }

  getValues () {
    return this.sheet.activeRange.getValues();
  }

  setValue (data) {
    this.sheet.activeRange.setValue(data);

    return this;
  }

  setValues (data) {
    this.sheet.activeRange.setValues(data);

    return this;
  }

  setFormat (format) {
    // https://developers.google.com/sheets/api/guides/formats
    this.sheet.activeRange.setNumberFormat(format);

    return this;
  }

  sort (column) {
    this.sheet.sort(column);

    return this;
  }

  sortMultiple (config) {
    this.sheet.activeRange.sort(config);

    return this;
  }

  setAlignHorizontal (align) {
    this.sheet.activeRange.setHorizontalAlignment(align);

    return this;
  }

  setAlignVertical (align) {
    this.sheet.activeRange.setVerticalAlignment(align);

    return this;
  }

  setWidth (column, width) {
    this.sheet.setColumnWidth(column, width);

    return this;
  }

  setWrap (wrap) {
    this.sheet.activeRange.setWrap(wrap);
  }

  setRowHeight (row, height) {
    this.sheet.setRowHeight(row, height);
  }

  setRowHeights (row, lastRow, height) {
    this.sheet.setRowHeights(row, lastRow, height);
  }

  setBorder (border) {
    this.sheet.activeRange.setBorder(border);
  }

  setHeader (config) {
    const range = this.sheet.getRange('A1:1');

    config.freeze && this.sheet.setFrozenRows(1);
    config.background && range.setBackground(config.background);
    config.color && range.setFontColor(config.color);
  }

  setBackground (color) {
    this.sheet.activeRange.setBackground(color);
  }

  setColor (color) {
    this.sheet.activeRange.setFontColor(color);
  }

  setWeight (weight) {
    this.sheet.activeRange.setFontWeight(weight);
  }

  setSize (size) {
    this.sheet.activeRange.setFontSize(size);
  }

  setFamily (family) {
    this.sheet.activeRange.setFontFamily(family);
  }

  setStyle (data) {
    data.background && this.setBackground(data.background);
    data.color && this.setColor(data.color);
  }

  expand (columnCount, rowCount) {
    if (columnCount > 0) {
      this.addColumns(columnCount);
    }
    if (rowCount > 0) {
      this.addRows(rowCount);
    }
    return this;
  }

  clearFormat () {
    this.sheet.clearFormats();
    return this;
  }

  copyTo (startRow, startColumn) {
    this.sheet.activeRange.copyTo(this.sheet.getRange(startRow, startColumn));

    return this;
  }

  copyToOnlyData (startRow, startColumn) {
    this.sheet.activeRange
      .copyTo(this.sheet.getRange(startRow, startColumn), { contentsOnly: true });

    return this;
  }

  copyToExt (originalSheet, targetSheet, startRow, startColumn) {
    this.active(targetSheet).getRange(startRow, startColumn);
    this.sheet.activeRange.copyTo(this.sheet.activeRange);
    return this.active(originalSheet);
  }

  addRows (count) {
    if (count < 1) {
      throw '추가할 행의 수가 잘못되었습니다.';
    }

    try {
      this.sheet.insertRowsAfter(this.sheet.getMaxRows(), count);
    } catch (e) {
      throw '행 추가 도중 문제가 발생했습니다.';
    }

    return this;
  }

  addColumns (count) {
    if (count < 1) {
      throw '추가할 열의 수가 잘못되었습니다.';
    }

    try {
      this.sheet.insertColumnsAfter(this.sheet.getMaxColumns(), count);
    } catch (e) {
      throw '열 추가 도중 문제가 발생했습니다.';
    }

    return this;
  }

  insertColumnAfter (index) {
    this.sheet.insertColumnAfter(index);

    return this;
  }

  insertRowAfter (index) {
    this.sheet.insertRowAfter(index);

    return this;
  }

  insertLastColumn (data, forceOneLength = false) {
    if (typeof data === 'object') {
      try {
        const dataLength = data.length;
        if (dataLength < 1) {
          throw '추가하려는 데이터가 없습니다.';
        }
        this.addColumns(forceOneLength ? 1 : dataLength)
          .sheet
          .getRange(1, this.sheet.getLastColumn() + 1, dataLength, 1)
          .setValues(data);
      } catch (e) {
        throw '추가하려는 셀 데이터가 실제 시트의 행 개수 보다 많습니다.';
      }
    } else if (typeof data === 'string') {
      try {
        this.sheet.getRange(1, this.sheet.getLastColumn() + 1).setValue(data);
      } catch (e) {
        throw '셀 데이터 입력 도중 문제가 발생했습니다.';
      }
    }

    return this;
  }

  insertLastRow (data) {
    if (typeof data === 'object') {
      const dataLength = data.length;
      if (dataLength < 1) {
        return this;
      }

      this.getRange(this.sheet.getLastRow() + 1, 1, dataLength, data[0].length).setValues(data);
    } else if (typeof data === 'string') {
      this.getRange(this.sheet.getLastRow() + 1, 1).setValue(data);
    } else {
      throw '올바른 데이터 형식이 아닙니다.';
    }

    return this;
  }

  merge () {
    this.sheet.activeRange.merge();

    return this;
  }

  mergeAcross () {
    this.sheet.activeRange.mergeAcross();

    return this;
  }

  clearAll () {
    this.sheet.clearContents();

    return this;
  }

  deleteColumn (index) {
    this.sheet.deleteColumn(index);

    return this;
  }

  deleteRow (index) {
    this.sheet.deleteRow(index + 1);

    return this;
  }

  minify (isRow) {
    if (!this.sheet) {
      return this;
    }

    let last = isRow ? this.sheet.getLastRow() : this.sheet.getLastColumn();
    const max = isRow ? this.sheet.getMaxRows() : this.sheet.getMaxColumns();

    if (isRow) {
      // 헤더 행을 포함
      const frozenRows = this.getFrozenRows();

      if (frozenRows > 0) {
        last += frozenRows;
      }
    }

    const avail = max - last;

    if (last < 1) {
      last = 1;
    }

    if (last !== max && avail > 1 && last > 0) {
      if (isRow) {
        this.sheet.deleteRows(last + 1, last === 1 ? avail - 1 : avail);
      } else {
        this.sheet.deleteColumns(last + 1, last === 1 ? avail - 1 : avail);
      }
    }

    return this;
  }

  minifyRows () {
    this.minify(true);

    return this;
  }

  minifyColumns () {
    this.minify(false);

    return this;
  }

  minifyAll () {
    this.minifyRows();
    this.minifyColumns();

    return this;
  }

  resizeColumns () {
    for (let i = 1, columnLength = this.sheet.getMaxColumns(); i <= columnLength; i += 1) {
      this.sheet.autoResizeColumn(i);
      this.sheet.setColumnWidth(i, this.sheet.getColumnWidth(i) + 30);
    }

    return this;
  }

  convertColumnToString (position) {
    const columnCharacters = ['A', 'B', 'C', 'D', 'E', 'F', 'G',
      'H', 'I', 'J', 'K', 'L', 'M', 'N',
      'O', 'P', 'Q', 'R', 'S', 'T', 'U',
      'V', 'W', 'X', 'Y', 'Z'];

    for (let i = 1, columnLength = columnCharacters.length; i < columnLength; i += 1) {
      if (position === i) {
        return columnCharacters[i - 1];
      }
    }

    return this;
  }
}
