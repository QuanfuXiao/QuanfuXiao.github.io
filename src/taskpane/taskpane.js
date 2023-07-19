/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function run() {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    // Read the range address
    range.load("address");

    // // Update the fill color
    range.format.fill.color = "yellow";

    await context.sync();
  });
}