/* eslint-disable @typescript-eslint/no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office, OfficeRuntime */

import { initializeApp } from "firebase/app";
import { getAuth, onAuthStateChanged, signInWithPopup, GoogleAuthProvider } from "firebase/auth";
import { getFirestore, doc, getDoc } from "firebase/firestore";

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
const provider = new GoogleAuthProvider();

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

async function launchSync() {
  const token = await getStorageItem("firebaseToken");
  if (!token) {
    return;
  }

  await Excel.run(async (context) => {
    // Your sync logic here using the token

    await context.sync();
  });
}

async function checkAuthStatus() {
  console.log("Checking Firebase authentication...");

  onAuthStateChanged(auth, async (user) => {
    if (user) {
      console.log("User is logged in:", user.email);
      document.getElementById("user-data").innerText = `Logged in as: ${user.email}`;

      // Fetch and display the user's Asana API key
      const apiKey = await fetchApiKey(user.uid);
      if (apiKey) {
        document.getElementById("api-key").innerText = `Asana API Key: ${apiKey}`;
        await OfficeRuntime.storage.setItem("asanaApiKey", apiKey);
      }
    } else {
      console.log("User is not logged in.");
      document.getElementById("user-data").innerText = "Not logged in";
      promptUserLogin();
    }
  });
}

function promptUserLogin() {
  const loginBtn = document.createElement("button");
  loginBtn.innerText = "Log in with Google";
  loginBtn.onclick = loginWithGoogle;
  document.getElementById("auth-container").appendChild(loginBtn);
}

async function loginWithGoogle() {
  try {
    const result = await signInWithPopup(auth, provider);
    console.log("User logged in:", result.user.email);
    checkAuthStatus(); // Rerun authentication check
  } catch (error) {
    console.error("Login failed:", error);
  }
}

async function fetchApiKey(uid) {
  console.log("Fetching API key for user:", uid);
  try {
    const userDoc = await getDoc(doc(db, "users", uid));
    if (userDoc.exists()) {
      return userDoc.data().apiKey;
    } else {
      console.error("No API key found for user.");
      return null;
    }
  } catch (error) {
    console.error("Error fetching API key:", error);
    return null;
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
