# gas-sheetman

[![license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/jooy2/gas-sheetman/blob/master/LICENSE)
[![Followers](https://img.shields.io/github/followers/jooy2?style=social)](https://github.com/jooy2)
![Stars](https://img.shields.io/github/stars/jooy2/gas-sheetman?style=social)
![Line Count](https://img.shields.io/tokei/lines/github/jooy2/gas-sheetman)
![Repo Size](https://img.shields.io/github/repo-size/jooy2/gas-sheetman)

Easy to use **[Google Apps Script](https://script.google.com/)** [Spreadsheet API](https://developers.google.com/apps-script/reference/spreadsheet) (using Method chaining)

## Features
- Easy to use, data insertion and styling
- Handling of external Google Sheets files

## Installation

### Installation using `clasp`
The clasp package must be installed for automatic installation in `Google Drive`.

```shell
$ npm install -g @google/clasp
```

If you have set up the development environment in this project, you can install it through `npm i`, so you don't need to install the package globally.

You may need a `clasp` login before proceeding. Skip if you have already done this:

```shell
$ clasp login
```

Then run one of the following commands:

```shell
$ clasp create
```

or...

```shell
$ npm run create
```

Follow the `clasp`'s prompts to create a spreadsheet in the `Google Drive` top-level path. After you open that sheet, the script will be installed automatically.

### Manual Installation
Put the `SheetMan.js` file in project directory into your new **[Google Apps Script](https://script.google.com/)** project file and use it.

You can also use additional `App.js` and `Test.js` files if necessary.

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
Related descriptions are attached to the entire method of `SheetMan.js`, and the main method of the default API commonly used in Spreadsheet of **Google Apps Script** is overridden.

Not all Spreadsheet methods may be compatible. If there is a method you would like to request, please leave an issue or send a PR.

See: https://developers.google.com/apps-script/reference/spreadsheet

## License
Copyright © 2021 Jooy2 Released under the MIT license
