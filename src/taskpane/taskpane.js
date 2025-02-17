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
  console.log("Office is ready. Checking authentication status...");

  // Check authentication status on page load
  onAuthStateChanged(auth, (user) => {
    if (user) {
      console.log("User is logged in:", user.email);
      userDataElement.innerText = `${user.email}`;
      fetchApiKey(user.uid);
      authContainer.style.display = "none";
      accountInfo.style.display = "block";
      backBtn.style.display = "block";
      logoutButton.style.display = "block";
      backBtn.click();
    } else {
      console.log("User is not logged in.");
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

function handleSelectionChanged(event) {
  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    const selection = context.workbook.getSelectedRange();
    selection.load("address");
    await context.sync();

    // Extract just the row numbers
    const address = selection.address;
    const match = address.match(/\$?\D+(\d+):?\$?\D*(\d+)?/);

    let rows = "";
    if (match) {
      rows = match[2] ? `${match[1]}-${match[2]}` : match[1];
    }

    sheetName.innerHTML = `<b>Sheet:</b> ${sheet.name}`;
    selectedRows.innerHTML = `<b>Selected Rows:</b> ${rows}`;
  });
}

async function launchSync() {
  //   const token = await getStorageItem("firebaseToken");
  //   if (!token) {
  //     console.log("NO TOKEN FOUND");
  //     return;
  //   }

  console.log("BEGIN LAUNCH SYNC");
  await Excel.run(async (context) => {
    // 1. Get the full rows for the current selection.
    console.log("1. Getting the full rows for the current selection...");

    // Get selected range
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load("address, values"); // Load values before expansion
    await context.sync();
    console.log("Selected Range Address: ", selectedRange.address);
    console.log("Selected Range Values: ", selectedRange.values);

    // Expand to entire row
    const entireRows = selectedRange.getEntireRow();
    entireRows.load("address");
    await context.sync();

    console.log("Expanded Rows Address: ", entireRows.address);

    const startRow = entireRows.address.split("!")[1].split(":")[0];
    const endRow = entireRows.address.split("!")[1].split(":")[1];
    console.log("Start row: ", startRow);
    console.log("End row: ", endRow);

    // Loop over selected rows
    var rows = [];

    // CONTINUE HERE


    // for (let i = startRow; i <= endRow; i++) {
    const worksheets = context.workbook.worksheets.getActiveWorksheet();
    console.log("Worksheets: ", worksheets);
    const range = worksheets.getRange(13, 1, 1, 6);
    console.log("Range: ", range);
    range.load("values");
    await context.sync();
    rows.push(range.values);
    // }

    console.log("Rows: ", rows);

    // 2. Open the first dialog to display the selected rows.
    console.log("2. Open the first dialog to display the selected rows.");
    const rowsData = encodeURIComponent(JSON.stringify(rows));
    const rowsDialogUrl = `./dialogs/rowsDialog.html?rows=${rowsData}`;
    Office.context.ui.displayDialogAsync(rowsDialogUrl, { height: 50, width: 50 }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const rowsDialog = result.value;

        // Send the rows data to the dialog
        rowsDialog.messageParent(JSON.stringify({ type: "displayRows", data: rows }));

        // Listen for messages from the rows dialog
        rowsDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          const message = JSON.parse(arg.message);
          if (message.type === "preview") {
            // 3. Build preview data (only for rows with a valid task)
            const previewData = rows.reduce((acc, row) => {
              // Assuming columns: 0: project, 1: doctor, 2: date, 3: task, 4: time, 5: comment
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

            // Open the preview dialog (ensure URL points to your previewDialog.html)
            const previewDataEncode = encodeURIComponent(JSON.stringify(previewData));
            const previewDialogUrl = `./dialogs/previewDialog.html?previewData=${previewDataEncode}`;
            Office.context.ui.displayDialogAsync(previewDialogUrl, { height: 50, width: 50 }, (previewResult) => {
              if (previewResult.status === Office.AsyncResultStatus.Succeeded) {
                const previewDialog = previewResult.value;

                // Send the preview data to the preview dialog
                previewDialog.messageParent(JSON.stringify({ type: "displayPreview", data: previewData }));

                // Handle messages from the preview dialog
                previewDialog.addEventHandler(Office.EventType.DialogMessageReceived, (e) => {
                  const previewMsg = JSON.parse(e.message);
                  if (previewMsg.type === "back") {
                    // Back: close preview and go back to rows dialog
                    previewDialog.close();
                  } else if (previewMsg.type === "launchSync") {
                    // 4. Launch sync: call handleSync and close both dialogs
                    handleSync(rows);
                    previewDialog.close();
                    rowsDialog.close();
                  }
                });
              }
            });
          } else if (message.type === "back") {
            // Back button clicked on the rows dialog: just close it.
            rowsDialog.close();
          }
        });
      }
    });
  });
}

async function handleSync(rows) {
  console.log("handleSync called with rows:", rows);
}

// Function to log in the user
loginButton.addEventListener("click", async () => {
  const email = emailField.value;
  const password = passwordField.value;

  try {
    const userCredential = await signInWithEmailAndPassword(auth, email, password);
    console.log("User logged in:", userCredential.user.email);
    fetchApiKey(userCredential.user.uid);
  } catch (error) {
    console.error("Login failed:", error.message);
    userDataElement.innerText = "Login failed. Check credentials.";
  }
});

// Function to fetch Asana API Key
async function fetchApiKey(uid) {
  console.log("Fetching API key for user:", uid);
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
    console.log("User logged out.");
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
