const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('🧾 Spreadit Tools')
    .addItem('Run Test', 'doTest')
    .addSeparator()
    .addItem('About', 'onAbout')
    .addToUi();
};

const onAbout = () => {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Copyright © Jooy2 Released under the MIT license.');
};
