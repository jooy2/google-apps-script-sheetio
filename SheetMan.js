/*
* SheetMan (https://github.com/jooy2/gas-sheetman)
* */
class SheetMan {
  constructor () {
    this.activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.originSheet = this.activeSheet;
    this.sheet = null;
  }

  createFile (title, nameForFirstSheet) {
    this.sheet = SpreadsheetApp.create(title);

    if (arguments.length === 2) {
      this.sheet.getActiveSheet().setName(nameForFirstSheet);
    }

    this.sheet.createdSheetId = this.getFileIdByUrl();

    return this;
  }

  createFileUsingApi (title) {
    const sheet = Sheets.newSpreadsheet();
    sheet.properties = Sheets.newSpreadsheetProperties();
    sheet.properties.title = title;

    const newFile = Sheets.SpreadSheets.create(sheet);
    this.sheet.createdSheetId = newFile.spreadsheetId;

    return this;
  }

  getFileIdByUrl () {
    // https://docs.google.com/spreadsheets/d/{SheetId}/edit#gid={gid}
    return this.sheet.getUrl().split('/')[5];
  }

  getFileId () {
    return this.sheet.createdSheetId
      ? this.sheet.createdSheetId
      : SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  getSheetId () {
    return this.sheet.getSheetId();
  }

  getSheetName () {
    return this.sheet.getName();
  }

  targetTo (sheetId) {
    if (sheetId) {
      this.isExternalSheet = true;
      this.sheet = SpreadsheetApp.openById(sheetId);
    } else {
      // Target to self sheet
      this.isExternalSheet = false;
      this.sheet = this.originSheet;
    }

    this.activeSheet = this.sheet;

    return this;
  }

  targetSelf () {
    this.targetTo();

    return this;
  }

  active (sheetName) {
    this.sheet = this.activeSheet.getSheetByName(sheetName);
    if (this.sheet) this.sheet.name = sheetName;

    if (sheetName.length < 1 || !this.sheet) {
      throw `'${sheetName}' The sheet does not exist. To create it, use create().`;
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
      this.sheet = this.activeSheet.insertSheet(sheetName, this.getSheetCount());
    } catch (e) {
      // Do nothing
    }
    return this;
  }

  rename (renameTo) {
    if (!renameTo || renameTo.length < 1) {
      throw `Could not change sheet '${this.sheet.name}' to '${renameTo}'.`;
    }

    this.sheet.setName(renameTo);

    return this;
  }

  destroy () {
    try {
      SpreadsheetApp.getActive().deleteSheet(this.sheet);
    } catch (e) {
      throw `Failed to delete sheet '${this.sheet.name}'.`;
    }

    return this;
  }

  destroyByName (sheetName) {
    if (sheetName && this.isExist(sheetName)) {
      return this.active(sheetName).destroy();
    }
    return this;
  }

  flush () {
    SpreadsheetApp.flush();

    return this;
  }

  moveActiveSheet (pos) {
    if (!pos || pos < 1) {
      throw 'Invalid pos value.';
    }
    this.sheet.moveActiveSheet();

    return this;
  }

  getId () {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
  }

  getActiveSheet () {
    return this.activeSheet;
  }

  getSheets () {
    if (!this.activeSheet) {
      return [];
    }
    return this.activeSheet.getSheets();
  }

  getSheetCount () {
    return this.getSheets().length;
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

    return this;
  }

  setRowHeight (row, height) {
    this.sheet.setRowHeight(row, height);
  }

  setRowHeights (row, lastRow, height) {
    this.sheet.setRowHeights(row, lastRow, height);
  }

  setBorder (border) {
    this.sheet.activeRange.setBorder(border);

    return this;
  }

  setHeader (config) {
    const range = this.sheet.getRange('A1:1');

    config.freeze && this.sheet.setFrozenRows(1);
    config.background && range.setBackground(config.background);
    config.color && range.setFontColor(config.color);
  }

  setBackground (color) {
    this.sheet.activeRange.setBackground(color);

    return this;
  }

  setColor (color) {
    this.sheet.activeRange.setFontColor(color);

    return this;
  }

  setWeight (weight) {
    this.sheet.activeRange.setFontWeight(weight);

    return this;
  }

  setSize (size) {
    this.sheet.activeRange.setFontSize(size);

    return this;
  }

  setFamily (family) {
    this.sheet.activeRange.setFontFamily(family);

    return this;
  }

  setStyle (data) {
    data.background && this.setBackground(data.background);
    data.color && this.setColor(data.color);

    return this;
  }

  setShowHyperlink (showHyperlink) {
    this.sheet.activeRange.setShowHyperlink(showHyperlink);

    return this;
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
    this.sheet.activeRange.clearFormat();

    return this;
  }

  clearFormats () {
    this.sheet.clearFormats();

    return this;
  }

  copyTo (destination) {
    if (this.sheet.activeRange) {
      this.sheet.activeRange.copyTo(destination);
    } else {
      this.sheet.copyTo(destination);
    }

    return this;
  }

  addRows (count) {
    if (count < 1) {
      throw 'Invalid number of rows to add.';
    }

    try {
      this.sheet.insertRowsAfter(this.sheet.getMaxRows(), count);
    } catch (e) {
      throw 'A problem occurred while adding rows.';
    }

    return this;
  }

  addColumns (count) {
    if (count < 1) {
      throw 'Invalid number of columns to add.';
    }

    try {
      this.sheet.insertColumnsAfter(this.sheet.getMaxColumns(), count);
    } catch (e) {
      throw 'A problem occurred while adding columns.';
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
      const dataLength = data.length;
      if (dataLength < 1) {
        throw 'There is no data to add.';
      }
      try {
        this.addColumns(forceOneLength ? 1 : dataLength)
          .sheet
          .getRange(1, this.sheet.getLastColumn() + 1, dataLength, 1)
          .setValues(data);
      } catch (e) {
        throw 'The cell data you want to add is more than the number of rows in the actual sheet.';
      }
    } else if (typeof data === 'string') {
      try {
        this.sheet.getRange(1, this.sheet.getLastColumn() + 1).setValue(data);
      } catch (e) {
        throw 'A problem occurred while entering cell data.';
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
      throw 'It is not a valid data type.';
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
    const columnCharacters = [
      'A', 'B', 'C', 'D', 'E', 'F', 'G',
      'H', 'I', 'J', 'K', 'L', 'M', 'N',
      'O', 'P', 'Q', 'R', 'S', 'T', 'U',
      'V', 'W', 'X', 'Y', 'Z',
    ];

    for (let i = 1, columnLength = columnCharacters.length; i < columnLength; i += 1) {
      if (position === i) {
        return columnCharacters[i - 1];
      }
    }

    return this;
  }
}
