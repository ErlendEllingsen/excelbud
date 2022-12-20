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

Office.actions.associate("MAPCELLS", function() {
  const context = new Excel.RequestContext();
  const range = context.workbook.getSelectedRange();
  
  applyMapping(context.workbook.worksheets.getActiveWorksheet());

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
