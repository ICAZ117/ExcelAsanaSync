/* eslint-disable no-redeclare */
/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, OfficeRuntime */

import { initializeApp } from "firebase/app";
import { getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { getFirestore, doc, getDoc } from "firebase/firestore";
import { asanaAPISync } from "lrh-asanaapisync";

class Queue {
  constructor(indentLevel) {
    this.items = [];
    this.indentLevel = indentLevel || 0;
  }

  enqueue(element) {
    this.items.push(element);
  }

  dequeue() {
    return this.isEmpty() ? null : this.items.shift();
  }

  front() {
    return this.isEmpty() ? null : this.items[0];
  }

  isEmpty() {
    return this.items.length === 0;
  }

  getQueue() {
    return this.items;
  }

  get(index) {
    return this.items[index];
  }

  size() {
    return this.items.length;
  }

  print() {
    console.log(this.items.join(" <- "));
  }

  clear() {
    this.items = [];
  }
}

class PrintJob {
  constructor(msg = "", params = undefined) {
    this.msg = msg;
    this.params = params;
  }
}

// Your Firebase project configuration
const firebaseConfig = {
  apiKey: "AIzaSyBvBL5nTfjdu85awdDkGTS-HtlUvTLcD2U",
  authDomain: "lrh-codebook.firebaseapp.com",
  projectId: "lrh-codebook",
  storageBucket: "lrh-codebook.appspot.com",
  messagingSenderId: "19502263714",
  appId: "1:19502263714:web:563e622ef36866ca5d16fb",
  measurementId: "G-VE6JHR065F",
};

// Initialize Firebase
const firebaseApp = initializeApp(firebaseConfig);
const auth = getAuth(firebaseApp);
const db = getFirestore(firebaseApp);

const taskConversion = {
  "Meeting: Intake": "Discovery Phase",
  "Meeting: Methods/Ideas": "Protocol Development",
  Analysis: "Statsitical Analysis",
  Products: "Publication",
  "Review/Revise Package": "IRB Package Preparation Phase",
  SAP: "IRB Package Preparation Phase",
  DRR: "IRB Package Preparation Phase",
  "Prep Work": "Statistical Analysis",
};

// Get UI elements
// Nav
const navToggler = document.getElementById("nav-toggler");
const navHome = document.getElementById("nav-home");
const navAccount = document.getElementById("nav-account");

// Auth Page
const authPage = document.getElementById("auth-page");
const authContainer = document.getElementById("auth-container");
const accountInfo = document.getElementById("account-info");
const emailField = document.getElementById("email");
const passwordField = document.getElementById("password");
const loginButton = document.getElementById("login-btn");
const logoutButton = document.getElementById("logout-btn");
const userDataElement = document.getElementById("user-data");
const apiKeyElement = document.getElementById("api-key");
const syncButton = document.getElementById("launchSync");
const backBtn = document.getElementById("back-btn");

// Sync Page
const syncPage = document.getElementById("sync-page");
const sheetName = document.getElementById("sheet-name");
const selectedRows = document.getElementById("selected-rows");

Office.onReady((info) => {
  debug("Office is ready. Checking authentication status...");

  // Check authentication status on page load
  onAuthStateChanged(auth, (user) => {
    if (user) {
      debug("User is logged in:", user.email);
      userDataElement.innerText = `${user.email}`;
      fetchApiKey(user.uid);
      authContainer.style.display = "none";
      accountInfo.style.display = "block";
      backBtn.style.display = "block";
      logoutButton.style.display = "block";
      backBtn.click();
    } else {
      debug("User is not logged in.");
      userDataElement.innerText = "Not logged in!";
      apiKeyElement.innerText = "Fetching API key...";
      authContainer.style.display = "block";
      accountInfo.style.display = "none";
      backBtn.style.display = "none";
      logoutButton.style.display = "none";
    }
  });
  if (info.host === Office.HostType.Excel) {
    // Ensure the DOM is fully loaded before accessing elements
    navToggler.onclick = function () {
      // if style is display: none, set to block, else set to none
      if (document.getElementById("navbarNav").style.display === "none") {
        document.getElementById("navbarNav").style.display = "block";
      } else {
        document.getElementById("navbarNav").style.display = "none";
      }
    };
    navHome.onclick = function () {
      authPage.style.display = "none";
      syncPage.style.display = "block";
      navToggler.click();
    };
    navAccount.onclick = function () {
      authPage.style.display = "block";
      syncPage.style.display = "none";
      navToggler.click();
    };
    backBtn.onclick = () => {
      authPage.style.display = "none";
      syncPage.style.display = "block";
    };
    const launchSyncButton = document.getElementById("launchSync");
    if (launchSyncButton) {
      launchSyncButton.onclick = launchSync;
    }
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.style.display = "flex";
    }

    // Monitor function to constantly update sheet-name
    Excel.run(async (context) => {
      context.workbook.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
    });

    handleSelectionChanged();
  }
});

function handleSelectionChanged() {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const selection = context.workbook.getSelectedRange();
    selection.load("address");
    await context.sync();

    // Extract just the row numbers
    const address = selection.address;
    const match = address.split("!")[1].match(/\$?\D+(\d+):?\$?\D*(\d+)?/);

    // debug("Address", address);
    // debug("Match", match);
    // debug("\n\n");

    let rows = "";
    if (match) {
      rows = match[2] && match[2] != match[1] ? `${match[1]}-${match[2]}` : match[1];
    }

    sheetName.innerHTML = `<b>Sheet:</b> ${sheet.name}`;
    selectedRows.innerHTML = `<b>Selected Rows:</b> ${rows}`;
    return rows;
  });
}

// ██       █████  ██    ██ ███    ██  ██████ ██   ██     ███████ ██    ██ ███    ██  ██████
// ██      ██   ██ ██    ██ ████   ██ ██      ██   ██     ██       ██  ██  ████   ██ ██
// ██      ███████ ██    ██ ██ ██  ██ ██      ███████     ███████   ████   ██ ██  ██ ██
// ██      ██   ██ ██    ██ ██  ██ ██ ██      ██   ██          ██    ██    ██  ██ ██ ██
// ███████ ██   ██  ██████  ██   ████  ██████ ██   ██     ███████    ██    ██   ████  ██████
async function launchSync() {
  LOGGING_INDENT = 0;

  addLog("\n\nBEGIN LAUNCH SYNC", undefined, 1);
  try {
    await Excel.run(async (context) => {
      // 1. Get the full rows for the current selection.
      addLog("1. Getting the full rows for the current selection", undefined, 1);
      // ┌

      // Get selected range
      addLog("1.1. Getting selected range", undefined, 1);
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load("address, values"); // Load values before expansion
      await context.sync();
      addLog("Selected Range Address: ", selectedRange.address);
      addLog("Selected Range Values: ", selectedRange.values, -1);

      // Expand to entire row
      addLog("1.2. Expanding selection to full rows", undefined, 1);
      const entireRows = selectedRange.getEntireRow();
      entireRows.load("address");
      await context.sync();

      addLog("Expanded Rows Address: ", entireRows.address);

      const startRow = entireRows.address.split("!")[1].split(":")[0];
      const endRow = entireRows.address.split("!")[1].split(":")[1];
      addLog("Start row: ", startRow);
      addLog("End row: ", endRow, -1);

      // Loop over selected rows
      addLog("1.3. Looping over selected rows", undefined, 1);
      var rows = [];
      const worksheets = context.workbook.worksheets.getActiveWorksheet();
      addLog("Worksheets: ", worksheets);
      const range = worksheets.getRange(`A${startRow}:G${endRow}`);
      addLog("Range: ", range);
      range.load("values");
      await context.sync();
      rows.push(range.values);

      // After you have your rows data:
      addLog("Rows: ", rows, -2);

      // 2. Open the first dialog to display the selected rows.
      addLog("2. Open the first dialog to display the selected rows.", undefined, 1);

      // Open dialog with promise-based handling
      const rowsData = encodeURIComponent(JSON.stringify(rows));
      const rowsDialogUrl = `https://localhost:8080/dialogs/rowsDialog?rows=${rowsData}`;

      // Convert dialog opening to promise
      const rowsDialog = await new Promise((resolve, reject) => {
        Office.context.ui.displayDialogAsync(
          rowsDialogUrl,
          { height: 400, width: 600, displayInIframe: false },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              addLog("Rows dialog opened successfully.", undefined, 0);
              resolve(result.value);
            } else {
              addLog("Failed to open rows dialog.", undefined, 0);
              reject(new Error("Failed to open rows dialog"));
            }
          }
        );
      });

      // Setup dialog message handling with promises
      const dialogResult = await new Promise((resolve) => {
        // Send the rows data to the dialog
        // rowsDialog.messageParent(JSON.stringify({ type: "displayRows", data: rows }));

        rowsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          console.log("Received message from dialog:", arg.message);

          const message = JSON.parse(arg.message);

          if (message.type === "preview") {
            // Handle preview - open second dialog
            const previewData = createPreviewData(rows);
            openPreviewDialog(previewData, rowsDialog).then(resolve);
          } else if (message.type === "back") {
            rowsDialog.close();
            resolve({ action: "back" });
          }
        });

        // Handle dialog closed event
        rowsDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          resolve({ action: "closed" });
        });
      });

      // Handle result of dialog flow
      if (dialogResult.action === "sync") {
        await handleSync(rows);
      }
    });
  } catch (error) {
    addLog("Error in launchSync: " + error.message);
  }

  addLog("END LAUNCH SYNC");
  LOGGING_INDENT = 0;
  formatLoggingQueue();
  printLoggingQueue();
  loggingQueue.clear();
}

// Helper function to open preview dialog
async function openPreviewDialog(previewData, rowsDialog) {
  const previewDataEncode = encodeURIComponent(JSON.stringify(previewData));
  const previewDialogUrl = `https://localhost:8080/dialogs/previewDialog?previewData=${previewDataEncode}`;

  const previewDialog = await new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(previewDialogUrl, { height: 50, width: 50 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to open preview dialog"));
      }
    });
  });

  return new Promise((resolve) => {
    // Send data to preview dialog
    previewDialog.messageParent(JSON.stringify({ type: "displayPreview", data: previewData }));

    // Handle messages from preview dialog
    previewDialog.addEventHandler(Office.EventType.DialogMessageReceived, (e) => {
      const previewMsg = JSON.parse(e.message);

      if (previewMsg.type === "back") {
        previewDialog.close();
        resolve({ action: "back" });
      } else if (previewMsg.type === "launchSync") {
        previewDialog.close();
        rowsDialog.close();
        resolve({ action: "sync" });
      }
    });

    // Handle dialog closed event
    previewDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
      resolve({ action: "closed" });
    });
  });
}

// Helper to create preview data
function createPreviewData(rows) {
  return rows.reduce((acc, row) => {
    const excelTask = row[3];
    if (Object.prototype.hasOwnProperty.call(taskConversion, excelTask)) {
      const asanaTask = taskConversion[excelTask];
      const previewComment = `${row[2]} - ${row[5]}`;
      acc.push({
        project: row[0],
        task: asanaTask,
        comment: previewComment,
      });
    }
    return acc;
  }, []);
}

async function handleSync(rows) {
  debug("handleSync called with rows:", rows);
}

// Function to log in the user
loginButton.addEventListener("click", async () => {
  const email = emailField.value;
  const password = passwordField.value;

  try {
    const userCredential = await signInWithEmailAndPassword(auth, email, password);
    debug("User logged in:", userCredential.user.email);
    fetchApiKey(userCredential.user.uid);
  } catch (error) {
    console.error("Login failed:", error.message);
    userDataElement.innerText = "Login failed. Check credentials.";
  }
});

// Function to fetch Asana API Key
async function fetchApiKey(uid) {
  debug("Fetching API key for user:", uid);
  try {
    const userDoc = await getDoc(doc(db, "users", uid));
    if (userDoc.exists()) {
      const apiKey = userDoc.data().apiKey;
      apiKeyElement.innerText = `${apiKey}`;
      await OfficeRuntime.storage.setItem("asanaApiKey", apiKey);
      syncButton.style.display = "block"; // Show sync button after API key is fetched
    } else {
      apiKeyElement.innerText = "No API key found.";
    }
  } catch (error) {
    console.error("Error fetching API key:", error);
  }
}

// Function to log out
logoutButton.addEventListener("click", async () => {
  try {
    await signOut(auth);
    debug("User logged out.");
    userDataElement.innerText = "Not logged in!";
    apiKeyElement.innerText = "Fetching API key...";
    apiKeyElement.innerText = "";
    syncButton.style.display = "none";
  } catch (error) {
    console.error("Logout failed:", error);
  }
});

async function getStorageItem(key) {
  try {
    return await OfficeRuntime.storage.getItem(key);
  } catch (error) {
    console.error(`Error getting storage item ${key}:`, error);
    return null;
  }
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

var DEBUG_INDENT = 0;

async function debug(message, params, updateIndent) {
  if (DEBUG_INDENT > 0) {
    if (updateIndent && updateIndent <= 0) {
      message = "│   ".repeat(DEBUG_INDENT - 1) + "└── " + message;
    } else {
      message = "│   ".repeat(DEBUG_INDENT - 1) + "├── " + message;
    }
  }

  if (params) {
    console.log(message, params);
  } else {
    console.log(message);
  }

  if (updateIndent) {
    DEBUG_INDENT += updateIndent;
  }
}

const loggingQueue = new Queue();
var LOGGING_INDENT = 0;

async function addLog(message, params, updateIndent) {
  var printJob = new PrintJob(message, params);

  // If the logging queue is empty, create a new print level and printjob
  if (loggingQueue.isEmpty()) {
    var printLevel = new Queue(LOGGING_INDENT);
    printLevel.enqueue(printJob);
    loggingQueue.enqueue(printLevel);
  } else {
    // Add printjob to the last print level
    var printLevel = loggingQueue.get(loggingQueue.size() - 1);
    printLevel.enqueue(printJob);
  }

  // If necessary, update logging indent and add next print level
  if (updateIndent) {
    LOGGING_INDENT += updateIndent;
    var printLevel = new Queue(LOGGING_INDENT);
    loggingQueue.enqueue(printLevel);
  }
}

async function formatLoggingQueue() {
  var mainQueue = loggingQueue.getQueue();
  const activeColumns = new Set();

  // Loop backwards through main queue
  for (var i = mainQueue.length - 1; i >= 0; i--) {
    var subQueue = mainQueue[i];

    // Unmark any active columns that are greater than the current indent level
    activeColumns.forEach((activeColumn) => {
      if (activeColumn > subQueue.indentLevel) {
        activeColumns.delete(activeColumn);
      }
    });

    // If current indent is 0, we just continue
    // Else, we need to update the string indentation
    if (subQueue.indentLevel > 0) {
      // Construct indentation string
      var indent = "";

      for (var k = 1; k <= subQueue.indentLevel; k++) {
        // If k is the current indent level, add a junction
        if (k === subQueue.indentLevel) {
          continue;
        }
        // Else if the column is active, add a pipe
        else if (activeColumns.has(k)) {
          indent += "│   ";
        }
        // Else, add a space
        else {
          indent += "    ";
        }
      }

      // Loop through all printJobs in subqueue and update indentation
      for (var j = 0; j < subQueue.size() - 1; j++) {
        var printJob = subQueue.get(j);
        printJob.msg = indent + "├── " + printJob.msg;
      }

      // Update the last element
      var lastPrintJob = subQueue.get(subQueue.size() - 1);

      // If the last element is active, add a junction
      if (activeColumns.has(subQueue.indentLevel)) {
        lastPrintJob.msg = indent + "├── " + lastPrintJob.msg;
      } else {
        lastPrintJob.msg = indent + "└── " + lastPrintJob.msg;
      }

      // Mark the current indent level as active
      activeColumns.add(subQueue.indentLevel);
    }
  }
}

async function printLoggingQueue() {
  var mainQueue = loggingQueue.getQueue();
  for (var i = 0; i < mainQueue.length; i++) {
    var subQueue = mainQueue[i];
    for (var j = 0; j < subQueue.size(); j++) {
      var printJob = subQueue.get(j);

      if (printJob.params) {
        console.log(printJob.msg, printJob.params);
      } else {
        console.log(printJob.msg);
      }
    }
  }
}
