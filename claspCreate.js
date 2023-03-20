/* eslint-disable no-console */
const { existsSync, unlinkSync } = require('fs');
const { exec, spawn } = require('child_process');

const CLASP_JSON_PATH = '.clasp.json';
let spreadSheetUrl;

if (existsSync(CLASP_JSON_PATH)) {
  unlinkSync(CLASP_JSON_PATH);
}

const claspProcess = spawn(
  'clasp',
  ['create', '--type', 'sheets', '--title', '"New Sheet with Sheetio"'],
  { shell: true }
);

claspProcess.stdout.on('data', (data) => {
  if (data.toString().indexOf('drive.google.com/open?id=') !== -1) {
    spreadSheetUrl = data
      .toString()
      .split('Created new Google Sheet: ')?.[1]
      ?.replace(/\r?\n/g, '');
  }
  console.log(data.toString());
});

claspProcess.stderr.on('data', (data) => {
  console.log(data.toString());
});

claspProcess.on('exit', (code) => {
  if (code === 0) {
    exec('clasp push -f', (error, stdout, stderr) => {
      if (error) {
        console.log(`error: ${error.message}`);
        return;
      }
      console.log(stdout);
      console.error(stderr);

      if (spreadSheetUrl) {
        let openStr;
        switch (process.platform) {
          case 'darwin':
            openStr = 'open';
            break;
          case 'win32':
            openStr = 'start';
            break;
          default:
            openStr = 'xdg-open';
            break;
        }
        exec(`${openStr} ${spreadSheetUrl}`);
      }
    });
  }
});
