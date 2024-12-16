# 🔌 Google Apps Script Spreadit

> [![license](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/jooy2/spreadit/blob/main/LICENSE) [![Followers](https://img.shields.io/github/followers/jooy2?style=social)](https://github.com/jooy2) ![Stars](https://img.shields.io/github/stars/jooy2/google-apps-script-sheetio?style=social) ![Line Count](https://img.shields.io/tokei/lines/github/jooy2/google-apps-script-sheetio) ![Commit Count](https://img.shields.io/github/commit-activity/y/jooy2/google-apps-script-sheetio) ![Repo Size](https://img.shields.io/github/repo-size/jooy2/google-apps-script-sheetio)

Easy to use **[Google Apps Script](https://script.google.com)** [Spreadsheet API](https://developers.google.com/apps-script/reference/spreadsheet) (using Method chaining)

## Features

- ⚡️ Installing scripts into your project Super-Fast!
- ⚡️ Easy to use, data insertion and styling
- ⚡️ Handling of external Google Sheets files
- ⚡️ Support for a variety of utility methods to simplify sheet operations

## Installation

### Step 1. Install `clasp`

In order to properly use all configuration of `clasp`, you need to change the Enabled setting of Google Apps Script API to 'Enable' after logging in with your Google account on the next page: https://script.google.com/home/usersettings

The clasp package must be installed for automatic installation in `Google Drive`.

```shell
$ npm install -g @google/clasp
```

If you have set up the development environment in this project, you can install it through `npm i`, so you don't need to install the package globally.

You may need a `clasp` login before proceeding. Skip if you have already done this:

```shell
$ clasp login
```

### Step 2-A. Create Google SpreadSheet script in your workspace

You can create Google Sheets and script projects in Google Drive. We recommend running your own scripts with everything integrated.

#### [Method 1] Create automatically (recommended)

This npm script helps you quickly and easily push scripts to your new spreadsheet file using `clasp`.

```shell
$ npm run create
```

#### [Method 2] Create manually

```shell
$ clasp create
```

Follow the `clasp`'s prompts to create a spreadsheet in the `Google Drive` top-level path. After you open that sheet, the script will be installed automatically.

Make sure a `.clasp.json` file is created in your project root. After that, you can install the script into the document with the command below.

```shell
$ clasp push
```

### Step 2-B. Manual Installation

If you already have a Google Script project created, or if you want to add your own script files, follow the steps below.

Put the `Spreadit.js` file in project directory into your new **[Google Apps Script](https://script.google.com)** project file and use it.

You can also use additional `App.js` and `Test.js` files if necessary.

## Usage

```javascript
// Create a Spreadit instance.
const Sheet = new Spreadit();

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

## Using API / Methods

Related descriptions are attached to the entire method of `Spreadit.js`, and the main method of the default API commonly used in Spreadsheet of **Google Apps Script** is overridden.

Not all Spreadsheet methods may be compatible. If there is a method you would like to request, please leave an issue or send a PR.

See: https://developers.google.com/apps-script/reference/spreadsheet

## Contributing

Anyone can contribute to the project by reporting new issues or submitting a pull request. For more information, please see [CONTRIBUTING.md](CONTRIBUTING.md).

## License

Please see the [LICENSE](LICENSE) file for more information about project owners, usage rights, and more.
