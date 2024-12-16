const doTest = () => {
  // TODO Run Demo
  const Sheet = new Spreadit();

  const targetSheet = Sheet.active('Example');
  targetSheet.insertLastRow([['A', 'B', 'C', 'D', 'E']]);

  // Initialize sheet
  // Sheet.destroyByName('Example')
  // Sheet.create('Example');
};
