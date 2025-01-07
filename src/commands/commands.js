/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // Register the function with Office.
    Office.actions.associate("action", actionOutlook);
  } else if (info.host === Office.HostType.Excel) {
    // Register the function with Office.
    Office.actions.associate("action", actionExcel);
  } else if (info.host === Office.HostType.PowerPoint) {
    // Register the function with Office.
    Office.actions.associate("action", actionPowerPoint);
  } else if (info.host === Office.HostType.Word) {
    // Register the function with Office.
    Office.actions.associate("action", actionWord);
  }

  if (info.host !== Office.HostType.Outlook) {
    Office.actions.associate("SHOWTASKPANE", showTaskpane);
    Office.actions.associate("HIDETASKPANE", hideTaskpane);
  }
});

async function showTaskpane() {
  try {
    await Office.addin.showAsTaskpane();
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
}

async function hideTaskpane() {
  try {
    await Office.addin.hide();
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
}


function actionOutlook(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

async function actionExcel(event) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function actionWord(event) {
  try {
    await Word.run(async (context) => {
      const paragraph = context.document.body.insertParagraph(
        "Hello World",
        Word.InsertLocation.end
      );
      paragraph.font.color = "blue";
      await context.sync();
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function actionPowerPoint(event) {
  try {
    await PowerPoint.run(async (context) => {
      const options = { coercionType: Office.CoercionType.Text };
      await Office.context.document.setSelectedDataAsync(" ", options);
      await Office.context.document.setSelectedDataAsync(
        "Hello World!",
        options
      );
    });
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
