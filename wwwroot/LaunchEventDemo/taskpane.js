/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { externalLink, AddinName } from "./common/constants.js";


/**
 * The add-in settings object.
 * @type {Office.RoamingSettings}
 */
let addinSettings;

/**
 * The URL of the event logging API.
 * @type {string}
 */
let logEventAPI = "";

/**
 * The delay in seconds for the API call.
 * @type {number}
 */
let apiDelayInSeconds = 0;

/**
 * Whether or not to block send if calling the API fails (e.g. server not available)
 */
let blockOnAPIFail = false;


/**
 * The Office.initialize function that gets called when the Office.js library is loaded.
 */
Office.initialize = function () {

  // Initialize instance variables to access API objects.
  addinSettings = Office.context.roamingSettings;
  var settingsUpdated = false;

  var apiUrl = addinSettings.get("apiUrl");
  if (apiUrl) {
      logEventAPI = apiUrl;
  }
  else {
      logEventAPI = window.location.origin + "/TestAPI/LogEventDelayed";
      addinSettings.set("apiUrl", logEventAPI);
      settingsUpdated = true;
  }

  var apiDelay = addinSettings.get("apiDelay");
  if (apiDelay > 0) {
      apiDelayInSeconds = apiDelay;
  } else if (apiDelay==null) {
      apiDelayInSeconds = 0;
      addinSettings.set("apiDelay", apiDelayInSeconds);
      settingsUpdated = true;
  }

  settingsUpdated = settingsUpdated | InitialiseAddinOption("blockOnAPIFail");
  settingsUpdated = settingsUpdated | InitialiseAddinOption("obtainAppointmentId");
  settingsUpdated = settingsUpdated | InitialiseAddinOption("showEventsOnMessage");

  if (settingsUpdated) {
    addinSettings.saveAsync(null);
    console.log("Settings written");
  } else {
    console.log("Settings read");
  }


  showAddinSetting("blockOnAPIFail");
  var blockOnAPIFailCheckbox = document.getElementById("blockOnAPIFailCheckbox");
  blockOnAPIFailCheckbox.addEventListener("change", checkboxChanged);

  showAddinSetting("obtainAppointmentId");
  var obtainAppointmentIdCheckbox = document.getElementById("obtainAppointmentIdCheckbox");
  obtainAppointmentIdCheckbox.addEventListener("change", checkboxChanged);

  showAddinSetting("showEventsOnMessage");
  var showEventsOnMessageCheckbox = document.getElementById("showEventsOnMessageCheckbox");
  showEventsOnMessageCheckbox.addEventListener("change", checkboxChanged); 

  document.getElementById("apiUrlInput").value = apiUrl;

  var apiDelayInput = document.getElementById("apiDelayInput");
  apiDelayInput.value = apiDelay;
  apiDelayInput.onchange = UpdateApiDelay;

  document.getElementById("openLink").onclick = openFolderLocationInWeb; // Add the click event for the new button
  document.getElementById("openDialog").onclick = openOfficeDialog;
}

/**
 * Updates the delay for the API call.
 */
function UpdateApiDelay() {
    console.log("UpdateApiDelay called");
    var apiDelay = document.getElementById("apiDelayInput").value;
    if (apiDelay != apiDelayInSeconds) {
        apiDelayInSeconds = Number(apiDelay);
        addinSettings.set("apiDelay", apiDelayInSeconds);
        addinSettings.saveAsync(null);
    }
}

function InitialiseAddinOption(settingName) {
  var settingChanged = false;
  var settingValue = addinSettings.get(settingName);
  if (settingValue == null) {
    addinSettings.set(settingValue, false);
    settingChanged = true;
  }
  console.log(settingName + " value: " + settingValue);
  return settingChanged;
}

function updateBlockOnAPIFail() {
  console.log("updateBlockOnAPIFail called");
  var blockOnAPIFailCheckbox = document.getElementById("blockOnAPIFailCheckbox");
  blockOnAPIFail = blockOnAPIFailCheckbox.checked;
  console.log("blockOnAPIFail set: " + blockOnAPIFail);
  addinSettings.set("blockOnAPIFail", blockOnAPIFail);
  addinSettings.saveAsync(null);
}

function obtainAppointmentIdChange() {
  console.log("obtainAppointmentId called");
  var obtainAppointmentIdCheckbox = document.getElementById("obtainAppointmentIdCheckbox");
  obtainAppointmentId = obtainAppointmentIdCheckbox.checked;
  console.log("obtainAppointmentId set: " + obtainAppointmentId);
  addinSettings.set("obtainAppointmentId", obtainAppointmentId);
  addinSettings.saveAsync(null);
}

function showAddinSetting(settingName) {
  var checkboxName = settingName + "Checkbox";
  var checkbox = document.getElementById(checkboxName);
  if (checkbox == null) {
    console.log("Couldn't locate " + checkboxName);
    return;
  }

  var addinSettingValue = addinSettings.get(settingName);
  console.log(settingName + " read from add-in: " + addinSettingValue);

  console.log(checkboxName + " aria-checked value: " + checkbox.getAttribute("aria-checked"));

  if ((addinSettingValue == "true" || addinSettingValue == true) && !checkbox.checked) {
    console.log("Ticking " + checkboxName);
    checkbox.checked = true;
    checkbox.setAttribute("aria-checked", true)
  }
  else if ((addinSettingValue == "false" || addinSettingValue == false) && checkbox.checked) {
    console.log("Unticking " + checkboxName);
    checkbox.checked = false;
    checkbox.setAttribute("aria-checked", false)
  }
  else {
    console.log(checkboxName + " is displaying correct value");
  }
}

function applyCheckboxSetting(settingName) {
  var checkbox = document.getElementById(settingName + "Checkbox");
  console.log(checkbox);
  checkboxChecked = checkbox.checked;
  console.log(settingName + " set: " + checkboxChecked);
  addinSettings.set(settingName, checkboxChecked);
}

function checkboxChanged() {
  // Called when any checkbox is changed from the UI (we read all checkbox values and set them)
  console.log("checkboxChanged called - reading add-in settings from UI");

  applyCheckboxSetting("blockOnAPIFail");
  applyCheckboxSetting("obtainAppointmentId");
  applyCheckboxSetting("showEventsOnMessage");

  addinSettings.saveAsync(null);
}

const openFolderLocationInWeb = async () => {
  const userProfile = Office.context.mailbox.userProfile;
  console.log(`Hello ${userProfile.displayName}`);

  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(externalLink);
    console.log(`Opening ${externalLink} in OWA`);
  } else {
    window.open(externalLink, "_blank", "noopener,noreferrer");
    console.log(`Opening ${externalLink} in Outlook classic`);
  }
};

const openOfficeDialog = async () => {
  let url = getAbsoluteURL(window.location.origin + window.location.pathname, "./dialog.html");

  const dialogOptions = { displayInIframe: false, height: 70, width: 50 };

  showDialog(url, dialogOptions, false).then((dialog) => {
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
      processMessage(args.message);
    });
  });
};

function processMessage(arg) {
  console.log(arg);
  console.log(`Message received from dialog`);

  if (arg.startsWith("http")) {
    // Open the external URL sent from the dialog
    window.open(arg, "_blank", "noopener,noreferrer");
  } else {
    console.log(`Unhandled message: ${arg}`);
  }
}

function openURLInBrowser() {
  // New function to open the external link based on the platform
  const externalLink = "https://microsoft.com";
  // Check if openBrowserWindow is available (indicating we're in OWA or New Outlook)
  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(externalLink);
    console.log(`Opening ${externalLink} in OWA`);
  } else {
    window.open(externalLink, "_blank", "noopener,noreferrer");
    console.log(`Opening ${externalLink} in Outlook classic`);
  }
}

function showDialog(url, dialogOptions, secondDialog) {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url, dialogOptions, async (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        if (secondDialog && asyncResult.error.code === 12007) {
          try {
            await this.sleep(1000);
            const res = await this.showDialog(url, dialogOptions, secondDialog);
            resolve(res);
          } catch (e) {
            reject(e);
          }
          // Recursive call
        } else {
          reject(asyncResult.error);
        }
      } else {
        resolve(asyncResult.value);
      }
    });
  });
}

function getAbsoluteURL(base, relative) {
  const stack = base.split("/");
  const parts = relative.split("/");
  stack.pop();

  for (let i = 0; i < parts.length; i += 1) {
    if (parts[i] === ".") {
      // Skip processing for '.'
    } else if (parts[i] === "..") {
      stack.pop();
    } else {
      stack.push(parts[i]);
    }
  }

  return stack.join("/");
}