/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    checkAuthStatus();
    document.getElementById("launchSync").onclick = () => tryCatch(launchSync);
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function launchSync() {
  const token = Office.context.roamingSettings.get("firebaseToken");
  if (!token) {
    openAuthDialog();
    return;
  }

  await Excel.run(async (context) => {
    // Your sync logic here using the token

    await context.sync();
  });
}

function checkAuthStatus() {
  const token = Office.context.roamingSettings.get("firebaseToken");
  if (token) {
    // User is authenticated
  } else {
    // User is not authenticated, open auth dialog
    openAuthDialog();
  }
}

function openAuthDialog() {
  Office.context.ui.displayDialogAsync('https://your-auth-url.com', { height: 50, width: 50 }, (result) => {
    const dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
      const token = args.message;
      Office.context.roamingSettings.set("firebaseToken", token);
      Office.context.roamingSettings.saveAsync();
      dialog.close();
    });
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
