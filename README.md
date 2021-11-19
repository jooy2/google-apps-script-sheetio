# gas-sheetman
Easy to use **[Google Apps Script](https://script.google.com/)** [Spreadsheet API](https://developers.google.com/apps-script/reference/spreadsheet) (using Method chaining)

## Features
- Easy to use, data insertion and styling
- Handling of external Google Sheets files

## Installation
Put the `SheetMan.js` file in project src directory into your new **[Google Apps Script](https://script.google.com/)** project file and use it.

## Usage

```javascript
// Create a SheetMan instance.
const Sheet = new SheetMan();

// Create a new sheet 'Users'.
Sheet.create('Users');

// Specifies where the currently active sheet is, allowing cells to be processed.
const targetSheet = Sheet.active('Users');

// The general method is similar to that of the Google Apps Script Spreadsheet.
targetSheet.getRange(1, 1).getValue();

// Continuous use possible with method chaining.
targetSheet
    .clearAll()
    .insertLastRow([['ID', 'Name', 'Age', 'Gender', 'Created', 'Updated', 'Subscription']])
    .insertLastRow([
        ['1', 'Lee', '26', 'M', '2021-11-01', '2021-11-01', 'Y'],
        ['2', 'James', '31', 'M', '2021-10-25', '2021-11-15', 'N'],
        ['3', 'Katy', '25', 'W', '2020-09-11', '2021-05-24', 'Y'],
        ['4', 'Betty', '27', 'W', '2021-03-09', '2021-11-11', 'N'],
        ['5', 'Mike', '32', 'M', '2020-07-29', '2020-09-06', 'N']
    ]);
```

## Methods
### Create and Edit File

❗❗❗ The document is not yet complete. Please refer to `src/SheetMan.js` for the full method. ❗❗❗

| Method | Params | Description | Example |
| --- | --- | --- | --- |
| `createFile` | <li>title **{String}**</li><li>nameForFirstSheet **{String}**</li> | Create a new Spreadsheet file. title is the name of the file, and nameForFirstSheet (optional) is the name of the sheet to be created when you first create it. | ```createFile()``` |

## License
Copyright © 2021 Jooy2 Released under the MIT license
