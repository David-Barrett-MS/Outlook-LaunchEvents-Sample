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

function FormatLog(data) {
    // Return log with add-in name and current time prepended
    var currentdate = new Date(); 
    var datetime = currentdate.getHours() + ":"  
                + currentdate.getMinutes() + ":" 
                + currentdate.getSeconds();
    return AddinName + " " + datetime + ": " + data;
}

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
    console.log(FormatLog("Settings written"));
  } else {
    console.log(FormatLog("Settings read"));
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
  document.getElementById("applyInsightMessage").onclick = applyInsightMessage;

  // Set up the ItemChanged event.
  if (Office.context.mailbox.item == null) {
    console.log(FormatLog("Item is null"));
  }
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
  console.log(FormatLog("ItemChanged event handler added"));
  updateTaskPaneUI(Office.context.mailbox.item);
}

/**
 * Updates the delay for the API call.
 */
function UpdateApiDelay() {
    console.log(FormatLog("UpdateApiDelay called"));
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
  return settingChanged;
}

function showAddinSetting(settingName) {
  var checkboxName = settingName + "Checkbox";
  var checkbox = document.getElementById(checkboxName);
  if (checkbox == null) {
    console.log(FormatLog("Couldn't locate " + checkboxName));
    return;
  }

  var checkboxLabel = document.getElementById(checkboxName + "Label");
  if (checkboxLabel == null) {
    console.log(FormatLog("Couldn't locate " + checkboxName + "Label"));
    return;
  }

  var addinSettingValue = addinSettings.get(settingName);
  console.log(FormatLog(settingName + " read from add-in: " + addinSettingValue));

  if ((addinSettingValue == "true" || addinSettingValue == true) && !checkbox.checked) {
    //console.log(FormatLog("Ticking " + checkboxName));
    checkbox.checked = true;
    checkboxLabel.classList.add("is-checked");
  }
  else if ((addinSettingValue == "false" || addinSettingValue == false) && checkbox.checked) {
    //console.log(FormatLog("Unticking " + checkboxName));
    checkbox.checked = false;
    checkbox.classList.remove("is-checked");
  }
  else {
    //console.log(FormatLog(checkboxName + " is displaying correct value"));
  }
}

function applyCheckboxSetting(settingName) {
  var checkbox = document.getElementById(settingName + "Checkbox");
  var checkboxChecked = checkbox.checked;
  console.log(FormatLog(settingName + " set: " + checkboxChecked));
  addinSettings.set(settingName, checkboxChecked);
}

export function checkboxChanged() {
  // Called when any checkbox is changed from the UI (we read all checkbox values and set them)
  console.log(FormatLog("checkboxChanged called - reading add-in settings from UI"));

  applyCheckboxSetting("blockOnAPIFail");
  applyCheckboxSetting("obtainAppointmentId");
  applyCheckboxSetting("showEventsOnMessage");

  addinSettings.saveAsync(null);
}

function openURL(linkToOpen) {
  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    console.log(FormatLog(`Opening ${linkToOpen} using openBrowserWindow`));
    Office.context.ui.openBrowserWindow(linkToOpen);
  } else {
    console.log(FormatLog(`Opening ${linkToOpen} using window.open`));
    window.open(linkToOpen, "_blank", "noopener,noreferrer");
  }  
}

const openFolderLocationInWeb = async () => {
  console.log(FormatLog("Open external link clicked"));
  openURL(externalLink);
};

const openOfficeDialog = async () => {
  let url = getAbsoluteURL(window.location.origin + window.location.pathname, "./dialog.html");

  const dialogOptions = { displayInIframe: true, height: 70, width: 50 };

  showDialog(url, dialogOptions, false).then((dialog) => {
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
      processMessage(args.message);
    });
  });
};

function processMessage(arg) {
  console.log(FormatLog(arg));
  console.log(FormatLog("Message received from dialog"));

  if (arg.startsWith("http")) {
    // Open the external URL sent from the dialog
    openURL(arg);
  } else {
    console.log(FormatLog(`Unhandled message: ${arg}`));
  }
}

function openURLInBrowser() {
  // New function to open the external link based on the platform
  const externalLink = "https://microsoft.com";
  // Check if openBrowserWindow is available (indicating we're in OWA or New Outlook)
  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    Office.context.ui.openBrowserWindow(externalLink);
    console.log(FormatLog(`Opening ${externalLink} in OWA`));
  } else {
    window.open(externalLink, "_blank", "noopener,noreferrer");
    console.log(FormatLog(`Opening ${externalLink} in Outlook classic`));
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


async function getInsightMessage() {
  return {
    type: Office.MailboxEnums.ItemNotificationMessageType.InsightMessage,
    message: "This is an InsightMessage",
    icon: "Icon.16x16",
    actions: [
      {
        actionText: "Process manually",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "msgComposeOpenPaneButton",
        contextData: "{}"
      }
    ]
  };
}

const applyInsightMessage = async () => {
  const notification = await getInsightMessage();

  console.log(FormatLog("Applying InsightMessage (from TaskPane button):", notification));
  Office.context.mailbox.item.notificationMessages.replaceAsync("InsightMessage", notification, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to apply InsightMessage:", asyncResult.error.message);
      return;
    }
    console.log(FormatLog("InsightMessage applied"));
  });  
}

// This function is called when the item changes in the task pane.
function itemChanged(eventArgs) {
  // Update UI based on the new current item.
  console.log(FormatLog("ItemChanged event fired"));
  updateTaskPaneUI(Office.context.mailbox.item);
}

/**
 * Write the current item's subject to the TaskPane
 */
function showSubject(subject) {
    var messageSubject = document.getElementById("messageSubjectInput");
    messageSubject.value = subject;
}

// This function updates the task pane UI based on the current item.
// All we actually do is write the subject to the console, but you could update the UI in other ways.
function updateTaskPaneUI(item) {
  if (item == null) {
    console.log(FormatLog("Item is null, unable to read subject"));
    return;
  }

  // Because we are using the same TaskPane in both compose and read modes, we need to check which mode we are in.
  // Easy test for this is to check the type of the subject property (it will only be a string in read mode).

  // Test for read mode
  if (typeof item.subject === "string") {
    // If the subject is a string, we are in read mode.
    console.log(FormatLog("Item subject (read mode): " + item.subject));
    showSubject(item.subject);
    console.log(FormatLog("Item recipients (read mode):"));
    const msgTo = item.to;
    for (let i = 0; i < msgTo.length; i++) {
      console.log(FormatLog(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")"));
    }    
    return;
  }
  
  // We are in compose mode
  item.subject.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            write(asyncResult.error.message);
            return;
        }
        console.log(FormatLog("Item subject (compose mode): " + asyncResult.value));
        showSubject(asyncResult.value);

        Office.context.mailbox.item.to.getAsync(function(asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const msgTo = asyncResult.value;
            console.log(FormatLog("Item recipients (compose mode):"));
            for (let i = 0; i < msgTo.length; i++) {
              console.log(FormatLog(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")"));
            }
          } else {
            console.error(asyncResult.error);
          }
        });        
      });    
}