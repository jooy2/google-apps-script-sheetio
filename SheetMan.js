/*
* SheetMan (https://github.com/jooy2/gas-sheetman)
* */
class SheetMan {
  constructor () {
    this.activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.originSheet = this.activeSheet;
    this.sheet = null;
  }

  /*
  * Begin of Google Apps Script Spreadsheet Wrapper
  * https://developers.google.com/apps-script/reference/spreadsheet
  * */

  getSheetId () {
    return this.sheet.getSheetId();
  }

  getName () {
    return this.sheet.getName();
  }

  destroy () {
    try {
      this.activeSheet.deleteSheet(this.sheet);
    } catch (e) {
      throw `Failed to delete sheet '${this.sheet.name}'.`;
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

  getSheets () {
    if (!this.activeSheet) {
      return [];
    }
    return this.activeSheet.getSheets();
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

  setNumberFormat (format) {
    // https://developers.google.com/sheets/api/guides/formats
    this.sheet.activeRange.setNumberFormat(format);

    return this;
  }

  sort (sortSpec) {
    if (this.sheet.activeRange) {
      this.sheet.activeRange.sort(sortSpec);
    } else {
      this.sheet.sort(sortSpec);
    }

    return this;
  }

  setHorizontalAlignment (align) {
    this.sheet.activeRange.setHorizontalAlignment(align);

    return this;
  }

  setVerticalAlignment (align) {
    this.sheet.activeRange.setVerticalAlignment(align);

    return this;
  }

  setColumnWidth (column, width) {
    this.sheet.setColumnWidth(column, width);

    return this;
  }

  setWrap (isWrapEnabled) {
    this.sheet.activeRange.setWrap(isWrapEnabled);

    return this;
  }

  setWraps (isWrapEnabled) {
    this.sheet.activeRange.setWraps(isWrapEnabled);

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

  setFontColor (color) {
    this.sheet.activeRange.setFontColor(color);

    return this;
  }

  setFontWeight (weight) {
    this.sheet.activeRange.setFontWeight(weight);

    return this;
  }

  setFontSize (size) {
    this.sheet.activeRange.setFontSize(size);

    return this;
  }

  setFontFamily (fontFamily) {
    this.sheet.activeRange.setFontFamily(fontFamily);

    return this;
  }

  setFontFamilies (fontFamilies) {
    this.sheet.activeRange.setFontFamilies(fontFamilies);

    return this;
  }

  setFontLine (fontLine) {
    this.sheet.activeRange.setFontLine(fontLine);

    return this;
  }

  setFontLines (fontLine) {
    this.sheet.activeRange.setFontLines(fontLine);

    return this;
  }

  setFormula (formula) {
    this.sheet.activeRange.setFormula(formula);

    return this;
  }

  setFormulas (formulas) {
    this.sheet.activeRange.setFormulas(formulas);

    return this;
  }

  setRichTextValue (value) {
    this.sheet.activeRange.setRichTextValue(value);

    return this;
  }

  setRichTextValues (values) {
    this.sheet.activeRange.setRichTextValues(values);

    return this;
  }

  setTextDirection (direction) {
    this.sheet.activeRange.setTextDirection(direction);

    return this;
  }

  setTextDirections (directions) {
    this.sheet.activeRange.setTextDirections(directions);

    return this;
  }

  setTextRotation (degrees) {
    this.sheet.activeRange.setTextRotation(degrees);

    return this;
  }

  setTextRotations (rotation) {
    this.sheet.activeRange.setTextRotations(rotation);

    return this;
  }

  setTextStyle (style) {
    this.sheet.activeRange.setTextStyle(style);

    return this;
  }

  setTextStyles (styles) {
    this.sheet.activeRange.setTextStyles(styles);

    return this;
  }

  setShowHyperlink (showHyperlink) {
    this.sheet.activeRange.setShowHyperlink(showHyperlink);

    return this;
  }

  setNote (note) {
    this.sheet.activeRange.setNote(note);

    return this;
  }

  setNotes (notes) {
    this.sheet.activeRange.setNotes(notes);

    return this;
  }

  isBlank () {
    return this.sheet.activeRange.isBlank();
  }

  isChecked () {
    return this.sheet.activeRange.isChecked();
  }

  isEndColumnBounded () {
    return this.sheet.activeRange.isEndColumnBounded();
  }

  isEndRowBounded () {
    return this.sheet.activeRange.isEndRowBounded();
  }

  isStartColumnBounded () {
    return this.sheet.activeRange.isStartColumnBounded();
  }

  isStartRowBounded () {
    return this.sheet.activeRange.isStartColumnBounded();
  }

  isPartOfMerge () {
    return this.sheet.activeRange.isPartOfMerge();
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

  moveTo (target) {
    this.sheet.activeRange.moveTo(target);

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

  mergeVertically () {
    this.sheet.activeRange.mergeVertically();

    return this;
  }

  clearContents () {
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

  check () {
    this.sheet.activeRange.check();
  }

  uncheck () {
    this.sheet.activeRange.uncheck();
  }

  /*
  * End of Google Apps Script Spreadsheet Wrapper
  * */

  /*
  * Begin of SheetMan methods
  * https://developers.google.com/apps-script/reference/spreadsheet
  * */

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
    this.sheet.activeRange = null;
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

  targetTo (sheetId) {
    if (sheetId) {
      this.isExternalSheet = true;
      this.sheet = SpreadsheetApp.openById(sheetId);
    } else {
      // Target to self sheet
      this.isExternalSheet = false;
      this.sheet = this.originSheet;
    }

    this.sheet.activeRange = null;
    this.activeSheet = this.sheet;

    return this;
  }

  targetSelf () {
    this.targetTo();

    return this;
  }

  active (sheetName) {
    this.sheet = this.activeSheet.getSheetByName(sheetName);

    if (this.sheet) {
      this.sheet.name = sheetName;
    }

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

  getActiveSheet () {
    return this.activeSheet;
  }

  getActiveRange () {
    return this.sheet.activeRange;
  }

  getSheetCount () {
    return this.getSheets().length;
  }

  removeRange () {
    this.sheet.activeRange = null;
  }

  create (sheetName) {
    try {
      this.sheet = this.activeSheet.insertSheet(sheetName, this.getSheetCount());

      if (this.sheet) {
        this.sheet.name = sheetName;
      }
    } catch (e) {
      this.active(sheetName);
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

  destroyByName (sheetName) {
    if (sheetName && this.isExist(sheetName)) {
      return this.active(sheetName).destroy();
    }
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

  insertColumnAfter (index) {
    this.sheet.insertColumnAfter(index);

    return this;
  }

  insertRowAfter (index) {
    this.sheet.insertRowAfter(index);

    return this;
  }

  minify (isRow) {
    if (!this.sheet) {
      return this;
    }

    let last = isRow ? this.sheet.getLastRow() : this.sheet.getLastColumn();
    const max = isRow ? this.sheet.getMaxRows() : this.sheet.getMaxColumns();

    if (isRow) {
      // Include header row
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

  static getColumnToString (position) {
    return [
      'A', 'B', 'C', 'D', 'E', 'F', 'G',
      'H', 'I', 'J', 'K', 'L', 'M', 'N',
      'O', 'P', 'Q', 'R', 'S', 'T', 'U',
      'V', 'W', 'X', 'Y', 'Z',
    ][position - 1];
  }

  /*
  * End of SheetMan methods
  * */
}
