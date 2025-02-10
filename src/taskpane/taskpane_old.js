/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, OfficeRuntime */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Ensure the DOM is fully loaded before accessing elements
    checkAuthStatus();
    const launchSyncButton = document.getElementById("launchSync");
    if (launchSyncButton) {
      launchSyncButton.onclick = () => tryCatch(launchSync);
    }
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.style.display = "flex";
    }
  }
});

var dialog = null;

async function launchSync() {
  const token = await getStorageItem("firebaseToken");
  if (!token) {
    openAuthDialog();
    return;
  }

  await Excel.run(async (context) => {
    // Your sync logic here using the token

    await context.sync();
  });
}

async function checkAuthStatus() {
  console.log("Checking auth status...");

  const token = await getStorageItem("firebaseToken");
  if (token) {
    console.log("Token found:", token);

    const userDataElement = document.getElementById("user-data");
    if (userDataElement) {
      userDataElement.innerText = `Token: ${token}`; // Display token in the DOM
    }
  } else {
    console.log("No token found, opening auth dialog...");
    openAuthDialog();
  }
}

function openAuthDialog() {
  Office.context.ui.displayDialogAsync(
    "https://lipsum.com/",
    { height: 50, width: 50, displayInIframe: false },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        dialog = asyncResult.value;
        // Listen for messages (i.e. the Firebase UID) from the dialog.
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        console.log("Dialog opened successfully.");

        // Inject fallback content in case the page isn't rendering
        dialog.messageChild("<h1 style='color:red;'>Hello from Excel!</h1>");
      } else {
        console.error("Dialog failed to open: " + asyncResult.error.message);
      }
    }
  );
}

function processMessage(arg) {
  console.log("Received message from dialog: " + arg);
  // Here we assume the auth page returns the Firebase UID as a plain string.
  // Display the UID in the 'result' div.
  document.getElementById("user-data").innerText = arg;
  // Optionally, close the dialog after receiving the message.
  if (dialog) {
    dialog.close();
  }
}

async function getStorageItem(key) {
  try {
    return await OfficeRuntime.storage.getItem(key);
  } catch (error) {
    console.error(`Error getting storage item ${key}:`, error);
    return null;
  }
}

// async function setStorageItem(key, value) {
//   try {
//     await OfficeRuntime.storage.setItem(key, value);
//   } catch (error) {
//     console.error(`Error setting storage item ${key}:`, error);
//   }
// }

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
