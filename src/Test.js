const onOpen = () => {
  SpreadsheetApp.getUi()
    .createMenu('✅ Run SheetMan Demo')
    .addItem('Run Demo', 'doGet')
    .addSeparator()
    .addItem('About', 'onAbout')
    .addToUi();
};

const doGet = () => {
  // TODO Run Demo
};

const onAbout = () => {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Copyright © 2021 Jooy2 Released under the MIT license.');
};
