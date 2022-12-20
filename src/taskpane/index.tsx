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

const title = "Contoso Task Pane Add-in";

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



      // console.log(usedRange.values);
      // usedRange.values.forEach(async (row, rowIndex) => {
      //   row.forEach(async (cell, colIndex) => {

      //     // get refernece to cell
      //     const cellRef = usedRange.getCell(rowIndex, colIndex);
      //     cellRef.load("address");
      //     await context.sync();



      //     console.log(`Cell ${rowIndex}, ${colIndex} has value ${cell} at address ${cellRef.address}`);
      //   });
      // });

      // /**
      //  * Insert your Excel code here
      //  */
      // const range = context.workbook.getSelectedRange();

      // // Read the range address
      // range.load("address");

      // // Update the fill color
      // range.format.fill.color = "yellow";

      // await context.sync();
      // console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }



  // const rangeFormat = range.format;
  // rangeFormat.fill.load();
  // const colors = ["#FFFFFF", "#C7CC7A", "#7560BA", "#9DD9D2", "#FFE1A8", "#E26D5C"];
  // return context.sync().then(function() {
  //   const rangeTarget = context.workbook.getSelectedRange();
  //   let currentColor = -1;
  //   for (let i = 0; i < colors.length; i++) {
  //     if (colors[i] == rangeFormat.fill.color) {
  //       currentColor = i;
  //       break;
  //     }
  //   }
  //   if (currentColor == -1) {
  //     currentColor = 0;
  //   } else if (currentColor == colors.length - 1) {
  //     currentColor = 0;
  //   } else {
  //     currentColor++;
  //   }
  //   rangeTarget.format.fill.color = colors[currentColor];
  //   return context.sync();
  // });
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
