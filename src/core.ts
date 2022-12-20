
export const applyMapping = async (sheet: Excel.Worksheet) => {
  const context = new Excel.RequestContext();
  // Iterate over each cell with content in the worksheet
  const usedRange = sheet.getUsedRange().load(["rowCount", "values"]);
  await context.sync();

  for (let row = 0; row < usedRange.rowCount; row++) {
    for (let col = 0; col < usedRange.values[row].length; col++) {
      const cell = usedRange.getCell(row, col);
      cell.load(["address", "formulas", "formulasR1C1", "values", "valueTypes"]);
      await context.sync();
      console.log(`Cell ${row}, ${col} has value ${usedRange.values[row][col]} at address ${cell.address}`);
      const cellValue = usedRange.values[row][col];

      // Skip empty cells
      if (cellValue === undefined || cellValue === null || cellValue === "") {
        continue;
      }

      // Determine if the range object (cell) references another cell
      const formulaContent = cell.formulas[0][0];
      const dataType = cell.valueTypes[0][0];

      // Skip string cells
      if (dataType === 'String') {
        continue;
      }

      const isNumeric = !isNaN(Number(formulaContent));
      if (isNumeric) {
        cell.format.fill.color = "yellow";
        // await context.sync();
        continue;
      }

      // Detect reference to another worksheet
      if ((formulaContent as string).includes('!')) {
        cell.format.fill.color = "cyan";
        // await context.sync();
        continue;
      }

      // Check for characters in the formula
      const isFormula = /[a-zA-Z]/.test(formulaContent);
      if (isFormula) {
        cell.format.fill.color = "red";
        // await context.sync();
        continue;
      }

      console.log(JSON.stringify(cell));

      console.log(formulaContent);

    }
  }

  await context.sync();
}