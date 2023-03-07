/*
 * SheetIO (@Jooy2, https://github.com/jooy2/google-apps-script-sheetio)
 * */
class SheetIO {
  constructor() {
    this.activeSheet = SpreadsheetApp.getActiveSpreadsheet();
    this.originSheet = this.activeSheet;
    this.sheet = null;
  }

  /* =================================================================
   * [BEGIN] Google Apps Script Spreadsheet Wrapper
   * https://developers.google.com/apps-script/reference/spreadsheet
   * ================================================================= */

  /* -----------------------------------------------------------------
   * [BEGIN] API::Sheet
   * ----------------------------------------------------------------- */

  getSheetId() {
    return this.sheet.getSheetId();
  }

  getName() {
    return this.sheet.getName();
  }

  destroy() {
    try {
      this.activeSheet.deleteSheet(this.sheet);
    } catch (e) {
      throw `Failed to delete sheet '${this.sheet.name}'.`;
    }

    return this;
  }

  flush() {
    SpreadsheetApp.flush();

    return this;
  }

  moveActiveSheet(pos) {
    if (!pos || pos < 1) {
      throw 'Invalid pos value.';
    }
    this.sheet.moveActiveSheet();

    return this;
  }

  getId() {
    return this.originSheet.getId();
  }

  getSheets() {
    if (!this.activeSheet) {
      return [];
    }
    return this.activeSheet.getSheets();
  }

  getMaxRows() {
    return this.sheet.getMaxRows();
  }

  getMaxColumns() {
    return this.sheet.getMaxColumns();
  }

  getFrozenRows() {
    return this.sheet.getFrozenRows();
  }

  getFrozenColumns() {
    return this.sheet.getFrozenColumns();
  }

  getColumnWidth(columnPosition) {
    return this.sheet.getColumnWidth(columnPosition);
  }

  getRange(startRow, startColumn, rows, columns) {
    if (arguments.length === 1) {
      this.sheet.activeRange = this.sheet.getRange(startRow);
    } else if (arguments.length === 2) {
      this.sheet.activeRange = this.sheet.getRange(startRow, startColumn);
    } else {
      this.sheet.activeRange = this.sheet.getRange(startRow, startColumn, rows, columns);
    }

    return this;
  }

  getDataRange() {
    this.sheet.activeRange = this.sheet.getDataRange();

    return this;
  }

  setColumnWidth(column, width) {
    this.sheet.setColumnWidth(column, width);

    return this;
  }

  setRowHeight(row, height) {
    this.sheet.setRowHeight(row, height);
  }

  setRowHeights(row, lastRow, height) {
    this.sheet.setRowHeights(row, lastRow, height);
  }

  setHeader(config) {
    const range = this.sheet.getRange('A1:1');

    config.freeze && this.sheet.setFrozenRows(1);
    config.background && range.setBackground(config.background);
    config.color && range.setFontColor(config.color);
  }

  autoResizeColumn(columnPosition) {
    this.sheet.autoResizeColumn(columnPosition);

    return this;
  }

  clearFormats() {
    this.sheet.clearFormats();

    return this;
  }

  clearNotes() {
    this.sheet.clearNotes();

    return this;
  }

  clearContents() {
    this.sheet.clearContents();

    return this;
  }

  deleteColumn(index) {
    this.sheet.deleteColumn(index);

    return this;
  }

  deleteRow(index) {
    this.sheet.deleteRow(index + 1);

    return this;
  }

  /* -----------------------------------------------------------------
   * [END] API::Sheet
   * ----------------------------------------------------------------- */

  /* -----------------------------------------------------------------
   * [BEGIN] API::Sheet::Range
   * ----------------------------------------------------------------- */

  getA1Notation() {
    return this.sheet.activeRange.getA1Notation();
  }

  getBackground() {
    return this.sheet.activeRange.getBackground();
  }

  getBackgroundObject() {
    return this.sheet.activeRange.getBackgroundObject();
  }

  getBackgroundObjects() {
    return this.sheet.activeRange.getBackgroundObjects();
  }

  getBackgrounds() {
    return this.sheet.activeRange.getBackgrounds();
  }

  getBandings() {
    return this.sheet.activeRange.getBandings();
  }

  getCell(row, column) {
    return this.sheet.activeRange.getCell(row, column);
  }

  getColumn() {
    return this.sheet.activeRange.getColumn();
  }

  getDataRegion(dimension) {
    if (dimension) {
      return this.sheet.activeRange.getDataRegion(dimension);
    }

    return this.sheet.activeRange.getDataRegion();
  }

  getDataSourceFormula() {
    return this.sheet.activeRange.getDataSourceFormula();
  }

  getDataSourceFormulas() {
    return this.sheet.activeRange.getDataSourceFormulas();
  }

  getDataSourcePivotTables() {
    return this.sheet.activeRange.getDataSourcePivotTables();
  }

  getDataSourceTables() {
    return this.sheet.activeRange.getDataSourceTables();
  }

  getDataSourceUrl() {
    return this.sheet.activeRange.getDataSourceUrl();
  }

  getDataTable(firstRowIsHeader) {
    if (firstRowIsHeader) {
      return this.sheet.activeRange.getDataTable(firstRowIsHeader);
    }

    return this.sheet.activeRange.getDataTable();
  }

  getDataValidation() {
    return this.sheet.activeRange.getDataValidation();
  }

  getDataValidations() {
    return this.sheet.activeRange.getDataValidations();
  }

  getDeveloperMetadata() {
    return this.sheet.activeRange.getDeveloperMetadata();
  }

  getDisplayValue() {
    return this.sheet.activeRange.getDisplayValue();
  }

  getDisplayValues() {
    return this.sheet.activeRange.getDisplayValues();
  }

  getFilter() {
    return this.sheet.activeRange.getFilter();
  }

  getFontColorObject() {
    return this.sheet.activeRange.getFontColorObject();
  }

  getFontColorObjects() {
    return this.sheet.activeRange.getFontColorObjects();
  }

  getFontFamilies() {
    return this.sheet.activeRange.getFontFamilies();
  }

  getFontFamily() {
    return this.sheet.activeRange.getFontFamily();
  }

  getFontLine() {
    return this.sheet.activeRange.getFontLine();
  }

  getFontLines() {
    return this.sheet.activeRange.getFontLines();
  }

  getFontSize() {
    return this.sheet.activeRange.getFontSize();
  }

  getFontSizes() {
    return this.sheet.activeRange.getFontSizes();
  }

  getFontStyle() {
    return this.sheet.activeRange.getFontStyle();
  }

  getFontStyles() {
    return this.sheet.activeRange.getFontStyles();
  }

  getFontWeight() {
    return this.sheet.activeRange.getFontWeight();
  }

  getFontWeights() {
    return this.sheet.activeRange.getFontWeights();
  }

  getFormula() {
    return this.sheet.activeRange.getFormula();
  }

  getFormulaR1C1() {
    return this.sheet.activeRange.getFormulaR1C1();
  }

  getFormulas() {
    return this.sheet.activeRange.getFormulas();
  }

  getFormulasR1C1() {
    return this.sheet.activeRange.getFormulasR1C1();
  }

  getGridId() {
    return this.sheet.activeRange.getGridId();
  }

  getWidth() {
    return this.sheet.activeRange.getWidth();
  }

  getHeight() {
    return this.sheet.activeRange.getHeight();
  }

  getHorizontalAlignment() {
    return this.sheet.activeRange.getHorizontalAlignment();
  }

  getHorizontalAlignments() {
    return this.sheet.activeRange.getHorizontalAlignments();
  }

  getVerticalAlignment() {
    return this.sheet.activeRange.getVerticalAlignment();
  }

  getVerticalAlignments() {
    return this.sheet.activeRange.getVerticalAlignments();
  }

  getLastRow() {
    const lastRow = this.sheet.getLastRow();

    return lastRow === 0 ? 1 : lastRow;
  }

  getLastColumn() {
    const lastColumn = this.sheet.getLastColumn();

    return lastColumn === 0 ? 1 : lastColumn;
  }

  getMergedRanges() {
    return this.sheet.activeRange.getMergedRanges();
  }

  getNextDataCell(direction) {
    return this.sheet.activeRange.getNextDataCell(direction);
  }

  getNote() {
    return this.sheet.activeRange.getNote();
  }

  getNotes() {
    return this.sheet.activeRange.getNotes();
  }

  getNumColumns() {
    return this.sheet.activeRange.getNumColumns();
  }

  getNumRows() {
    return this.sheet.activeRange.getNumRows();
  }

  getNumberFormat() {
    return this.sheet.activeRange.getNumberFormat();
  }

  getNumberFormats() {
    return this.sheet.activeRange.getNumberFormats();
  }

  getRichTextValue() {
    return this.sheet.activeRange.getRichTextValue();
  }

  getRichTextValues() {
    return this.sheet.activeRange.getRichTextValues();
  }

  getRow() {
    return this.sheet.activeRange.getRow();
  }

  getRowIndex() {
    return this.sheet.activeRange.getRowIndex();
  }

  getSheet() {
    return this.sheet.activeRange.getSheet();
  }

  getTextStyle() {
    return this.sheet.activeRange.getTextStyle();
  }

  getTextStyles() {
    return this.sheet.activeRange.getTextStyles();
  }

  getValue() {
    return this.sheet.activeRange.getValue();
  }

  getValues() {
    return this.sheet.activeRange.getValues();
  }

  getWrap() {
    return this.sheet.activeRange.getWrap();
  }

  getWrapStrategies() {
    return this.sheet.activeRange.getWrapStrategies();
  }

  getWrapStrategy() {
    return this.sheet.activeRange.getWrapStrategy();
  }

  getWraps() {
    return this.sheet.activeRange.getWraps();
  }

  setValue(data) {
    this.sheet.activeRange.setValue(data);

    return this;
  }

  setValues(data) {
    this.sheet.activeRange.setValues(data);

    return this;
  }

  setNumberFormat(numberFormat) {
    // https://developers.google.com/sheets/api/guides/formats
    this.sheet.activeRange.setNumberFormat(numberFormat);

    return this;
  }

  setNumberFormats(numberFormats) {
    this.sheet.activeRange.setNumberFormats(numberFormats);

    return this;
  }

  sort(sortSpec) {
    if (this.sheet.activeRange) {
      this.sheet.activeRange.sort(sortSpec);
    } else {
      this.sheet.sort(sortSpec);
    }

    return this;
  }

  setHorizontalAlignment(alignment) {
    this.sheet.activeRange.setHorizontalAlignment(alignment);

    return this;
  }

  setVerticalAlignment(alignment) {
    this.sheet.activeRange.setVerticalAlignment(alignment);

    return this;
  }

  setHorizontalAlignments(alignments) {
    this.sheet.activeRange.setHorizontalAlignments(alignments);

    return this;
  }

  setVerticalAlignments(alignments) {
    this.sheet.activeRange.setVerticalAlignments(alignments);

    return this;
  }

  setWrap(isWrapEnabled) {
    this.sheet.activeRange.setWrap(isWrapEnabled);

    return this;
  }

  setWraps(isWrapEnabled) {
    this.sheet.activeRange.setWraps(isWrapEnabled);

    return this;
  }

  setBorder(border) {
    this.sheet.activeRange.setBorder(border);

    return this;
  }

  setBackground(color) {
    this.sheet.activeRange.setBackground(color);

    return this;
  }

  setFontColor(color) {
    this.sheet.activeRange.setFontColor(color);

    return this;
  }

  setFontColorObject(color) {
    this.sheet.activeRange.setFontColorObject(color);

    return this;
  }

  setFontColorObjects(colors) {
    this.sheet.activeRange.setFontColorObjects(colors);

    return this;
  }

  setFontColors(colors) {
    this.sheet.activeRange.setFontColors(colors);

    return this;
  }

  setFontSize(size) {
    this.sheet.activeRange.setFontSize(size);

    return this;
  }

  setFontFamily(fontFamily) {
    this.sheet.activeRange.setFontFamily(fontFamily);

    return this;
  }

  setFontFamilies(fontFamilies) {
    this.sheet.activeRange.setFontFamilies(fontFamilies);

    return this;
  }

  setFontLine(fontLine) {
    this.sheet.activeRange.setFontLine(fontLine);

    return this;
  }

  setFontLines(fontLine) {
    this.sheet.activeRange.setFontLines(fontLine);

    return this;
  }

  setFontStyle(fontStyle) {
    this.sheet.activeRange.setFontStyle(fontStyle);

    return this;
  }

  setFontStyles(fontStyles) {
    this.sheet.activeRange.setFontStyles(fontStyles);

    return this;
  }

  setFontWeight(fontWeight) {
    this.sheet.activeRange.setFontWeight(fontWeight);

    return this;
  }

  setFontWeights(fontWeights) {
    this.sheet.activeRange.setFontWeights(fontWeights);

    return this;
  }

  setFormula(formula) {
    this.sheet.activeRange.setFormula(formula);

    return this;
  }

  setFormulas(formulas) {
    this.sheet.activeRange.setFormulas(formulas);

    return this;
  }

  setFormulaR1C1(formula) {
    this.sheet.activeRange.setFormulaR1C1(formula);

    return this;
  }

  setFormulasR1C1(formulas) {
    this.sheet.activeRange.setFormulasR1C1(formulas);

    return this;
  }

  setRichTextValue(value) {
    this.sheet.activeRange.setRichTextValue(value);

    return this;
  }

  setRichTextValues(values) {
    this.sheet.activeRange.setRichTextValues(values);

    return this;
  }

  setTextDirection(direction) {
    this.sheet.activeRange.setTextDirection(direction);

    return this;
  }

  setTextDirections(directions) {
    this.sheet.activeRange.setTextDirections(directions);

    return this;
  }

  setTextRotation(degrees) {
    this.sheet.activeRange.setTextRotation(degrees);

    return this;
  }

  setTextRotations(rotation) {
    this.sheet.activeRange.setTextRotations(rotation);

    return this;
  }

  setTextStyle(style) {
    this.sheet.activeRange.setTextStyle(style);

    return this;
  }

  setTextStyles(styles) {
    this.sheet.activeRange.setTextStyles(styles);

    return this;
  }

  setShowHyperlink(showHyperlink) {
    this.sheet.activeRange.setShowHyperlink(showHyperlink);

    return this;
  }

  setNote(note) {
    this.sheet.activeRange.setNote(note);

    return this;
  }

  setNotes(notes) {
    this.sheet.activeRange.setNotes(notes);

    return this;
  }

  setVerticalText(isVertical) {
    this.sheet.activeRange.setVerticalText(isVertical);

    return this;
  }

  isBlank() {
    return this.sheet.activeRange.isBlank();
  }

  isChecked() {
    return this.sheet.activeRange.isChecked();
  }

  isEndColumnBounded() {
    return this.sheet.activeRange.isEndColumnBounded();
  }

  isEndRowBounded() {
    return this.sheet.activeRange.isEndRowBounded();
  }

  isStartColumnBounded() {
    return this.sheet.activeRange.isStartColumnBounded();
  }

  isStartRowBounded() {
    return this.sheet.activeRange.isStartColumnBounded();
  }

  isPartOfMerge() {
    return this.sheet.activeRange.isPartOfMerge();
  }

  createDataSourcePivotTable(dataSource) {
    this.sheet.activeRange.createDataSourcePivotTable(dataSource);

    return this;
  }

  createDataSourceTable(dataSource) {
    this.sheet.activeRange.createDataSourceTable(dataSource);

    return this;
  }

  createDeveloperMetadataFinder() {
    this.sheet.activeRange.createDeveloperMetadataFinder();

    return this;
  }

  createFilter() {
    this.sheet.activeRange.createFilter();

    return this;
  }

  createPivotTable(sourceData) {
    this.sheet.activeRange.createPivotTable(sourceData);

    return this;
  }

  createTextFinder(findText) {
    this.sheet.activeRange.createTextFinder(findText);

    return this;
  }

  canEdit() {
    return this.sheet.activeRange.canEdit();
  }

  randomize() {
    this.sheet.activeRange = this.sheet.activeRange.randomize();

    return this;
  }

  removeCheckboxes() {
    this.sheet.activeRange = this.sheet.activeRange.removeCheckboxes();

    return this;
  }

  removeDuplicates() {
    this.sheet.activeRange = this.sheet.activeRange.removeDuplicates();

    return this;
  }

  trimWhitespace() {
    this.sheet.activeRange = this.sheet.activeRange.trimWhitespace();

    return this;
  }

  clearFormat() {
    this.sheet.activeRange.clearFormat();

    return this;
  }

  copyTo(destination) {
    if (this.sheet.activeRange) {
      this.sheet.activeRange.copyTo(destination);
    } else {
      this.sheet.copyTo(destination);
    }

    return this;
  }

  moveTo(target) {
    this.sheet.activeRange.moveTo(target);

    return this;
  }

  merge() {
    this.sheet.activeRange.merge();

    return this;
  }

  mergeAcross() {
    this.sheet.activeRange.mergeAcross();

    return this;
  }

  mergeVertically() {
    this.sheet.activeRange.mergeVertically();

    return this;
  }

  check() {
    this.sheet.activeRange.check();
  }

  uncheck() {
    this.sheet.activeRange.uncheck();
  }

  /* =================================================================
   * [END] Google Apps Script Spreadsheet Wrapper
   * ================================================================= */

  /* =================================================================
   * [BEGIN] Standalone utility methods used in SheetIO
   * See: https://github.com/jooy2/google-apps-script-sheetio/blob/master/README.md
   * ================================================================= */

  createFile(title, nameForFirstSheet) {
    this.sheet = SpreadsheetApp.create(title);

    if (arguments.length === 2) {
      this.sheet.getActiveSheet().setName(nameForFirstSheet);
    }

    this.sheet.createdSheetId = this.getFileIdByUrl();

    return this;
  }

  createFileUsingApi(title) {
    const sheet = Sheets.newSpreadsheet();
    sheet.properties = Sheets.newSpreadsheetProperties();
    sheet.properties.title = title;

    const newFile = Sheets.SpreadSheets.create(sheet);
    this.sheet.activeRange = null;
    this.sheet.createdSheetId = newFile.spreadsheetId;

    return this;
  }

  getFileIdByUrl() {
    // https://docs.google.com/spreadsheets/d/{SheetId}/edit#gid={gid}
    return this.sheet.getUrl().split('/')[5];
  }

  getFileId() {
    return this.sheet.createdSheetId ? this.sheet.createdSheetId : this.getId();
  }

  targetTo(sheetId) {
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

  targetSelf() {
    this.targetTo();

    return this;
  }

  active(sheetName) {
    this.sheet = this.activeSheet.getSheetByName(sheetName);

    if (this.sheet) {
      this.sheet.name = sheetName;
    }

    if (sheetName.length < 1 || !this.sheet) {
      throw `'${sheetName}' The sheet does not exist. To create it, use create().`;
    }

    return this;
  }

  isExist(sheetName) {
    if (this.isExternalSheet) {
      return this.sheet ? this.sheet.getSheetByName(sheetName) : null;
    }
    return this.activeSheet.getSheetByName(sheetName);
  }

  getActiveSheet() {
    return this.activeSheet;
  }

  getActiveRange() {
    return this.sheet.activeRange;
  }

  getSheetCount() {
    return this.getSheets().length;
  }

  removeRange() {
    this.sheet.activeRange = null;
  }

  create(sheetName) {
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

  rename(renameTo) {
    if (!renameTo || renameTo.length < 1) {
      throw `Could not change sheet '${this.sheet.name}' to '${renameTo}'.`;
    }

    this.sheet.setName(renameTo);

    return this;
  }

  destroyByName(sheetName) {
    if (sheetName && this.isExist(sheetName)) {
      return this.active(sheetName).destroy();
    }
    return this;
  }

  expand(columnCount, rowCount) {
    if (columnCount > 0) {
      this.addColumns(columnCount);
    }
    if (rowCount > 0) {
      this.addRows(rowCount);
    }
    return this;
  }

  addRows(count) {
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

  addColumns(count) {
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

  insertLastColumn(data, forceOneLength = false) {
    if (typeof data === 'object') {
      const dataLength = data.length;
      if (dataLength < 1) {
        throw 'There is no data to add.';
      }
      try {
        this.addColumns(forceOneLength ? 1 : dataLength)
          .sheet.getRange(1, this.sheet.getLastColumn() + 1, dataLength, 1)
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

  insertLastRow(data) {
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

  insertColumnAfter(index) {
    this.sheet.insertColumnAfter(index);

    return this;
  }

  insertRowAfter(index) {
    this.sheet.insertRowAfter(index);

    return this;
  }

  minify(isRow) {
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

  minifyRows() {
    this.minify(true);

    return this;
  }

  minifyColumns() {
    this.minify(false);

    return this;
  }

  minifyAll() {
    this.minifyRows();
    this.minifyColumns();

    return this;
  }

  resizeColumns() {
    for (let i = 1, columnLength = this.sheet.getMaxColumns(); i <= columnLength; i += 1) {
      this.sheet.autoResizeColumn(i);
      this.sheet.setColumnWidth(i, this.sheet.getColumnWidth(i) + 30);
    }

    return this;
  }

  static getColumnToString(position) {
    return [
      'A',
      'B',
      'C',
      'D',
      'E',
      'F',
      'G',
      'H',
      'I',
      'J',
      'K',
      'L',
      'M',
      'N',
      'O',
      'P',
      'Q',
      'R',
      'S',
      'T',
      'U',
      'V',
      'W',
      'X',
      'Y',
      'Z'
    ][position - 1];
  }

  /* =================================================================
   * [END] Standalone utility methods used in SheetIO
   * ================================================================= */
}
