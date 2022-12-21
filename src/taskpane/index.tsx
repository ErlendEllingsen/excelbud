import App from "./components/App";
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import { ThemeProvider } from "@fluentui/react";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { applyMapping } from "../core";

/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Excelbud Task Pane Add-in";

const modelColors = {
  'input': '#ffff71', // yellow
  'formula': '#f7c560', // darker yellow / orange
  'formula_other_sheet_dependent': 'cyan', // cyan
  'formula_other_file_dependent': 'd'
}

const render = (Component) => {
  ReactDOM.render(
    <AppContainer>
      <ThemeProvider>
        <Component title={title} isOfficeInitialized={isOfficeInitialized} />
      </ThemeProvider>
    </AppContainer>,
    document.getElementById("container")
  );
};

Office.actions.associate("MAPCELLS", async function () {
  const context = new Excel.RequestContext();
  try {
    await Excel.run(async (context) => {
      console.log('hi');
      const sheet = context.workbook.worksheets.getActiveWorksheet();

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

          console.log(JSON.stringify(cell));
          console.log(formulaContent);

          const isNumeric = !isNaN(Number(formulaContent));
          if (isNumeric) {
            cell.format.fill.color = "#ffff71"; // yellow
            // await context.sync();
            continue;
          }

          // Detect reference to another worksheet
          const cellRefRegex = /[a-zA-Z$]+[0-9$]+/;
          const isFormula = cellRefRegex.test(formulaContent);

          if ((formulaContent as string).includes('!') && isFormula) {
            cell.format.fill.color = "cyan";
            // await context.sync();
            continue;
          }

          // Check for characters in the formula
          if (isFormula) {
            cell.format.fill.color = "#f7c560"; // darker yellow / orange
            // await context.sync();
            continue;
          }



        }
      }

      await context.sync();

    });
  } catch (error) {
    console.error(error);
  }

});

Office.actions.associate("INVERTSIGN", async function () {

  const context = new Excel.RequestContext();
  try {

    await Excel.run(async (context) => {

      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Get active cells
      const range = context.workbook.getSelectedRange();
      range.load(["address", "rowCount", "values"]);
      await context.sync();

      for (let row = 0; row < range.rowCount; row++) {
        for (let col = 0; col < range.values[row].length; col++) {
          const cell = range.getCell(row, col);
          cell.load(["address", "formulas", "formulasR1C1", "values", "valueTypes"]);
          await context.sync();
          console.log(`Cell ${row}, ${col} has value ${range.values[row][col]} at address ${cell.address}`);
          const cellValue = range.values[row][col];

          // Skip empty cells
          if (cellValue === undefined || cellValue === null || cellValue === "") {
            continue;
          }

          // Determine if the range object (cell) references another cell
          const formulaContent = `${cell.formulas[0][0]}`;

          const wrappedNegativeRegex = /^=-1\*\(.*\)$/;
          const isWrappedNegative = wrappedNegativeRegex.test(formulaContent);
          console.log(isWrappedNegative)

          if (isWrappedNegative) {
            // Remove beginning and end of formula
            const unwrappedFormula = '=' + formulaContent.substring(5, formulaContent.length - 1);
            cell.formulas = [[unwrappedFormula]];
            await context.sync();
            continue;
          }

          const beganWithEquals = formulaContent.startsWith('=');
          const oldContent = beganWithEquals ? formulaContent.substring(1, formulaContent.length) : formulaContent;

          // Wrap with negative sign
          const wrappedFormula = `=-1*(${oldContent})`;
          cell.formulas = [[wrappedFormula]];
          await context.sync();


        }
      }
    });

  } catch (error) {
    console.error(error);
  }


});


/* Render application after Office initializes */
Office.onReady(() => {
  isOfficeInitialized = true;
  render(App);
});


if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
