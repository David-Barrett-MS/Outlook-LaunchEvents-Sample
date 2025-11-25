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
 * Test recipient (used when sending emails or invites during tests)
 * @type {string}
 */
let testRecipient = "";

var addInOptions=["blockOnAPIFail", "obtainAppointmentId", "showEventsOnMessage", "sendClientInfo", "showCustomSmartAlertDialog"];


function FormatLog(data) {
  // Return log with add-in name and current time prepended
  let currentTime = new Date().toLocaleTimeString('en-US', { hour12: false });
  return AddinName + " " + currentTime + ": " + data;
}

function ShowConsoleInTaskPane(container) {
  // Add console output at the bottom of the TaskPane

  if (container==null) {
    // Create the container if not supplied
    container = document.createElement('div');
    container.style.border = '1px solid #000';
    container.style.marginTop = '20px';
    document.body.appendChild(container);
  }

  // code element
  const outputDiv = document.createElement('code');
  outputDiv.style.whiteSpace = 'pre-wrap'; // Preserve formatting
  outputDiv.style.fontSize = '11px';
  outputDiv.style.wordBreak = 'break-word';
  container.appendChild(outputDiv);
  //document.body.appendChild(consoleDiv);

  // Save the original console.log function
  const originalConsoleLog = console.log;

  // Override console.log
  console.log = function (...args) {
    // Append the log message to the outputDiv
    const message = args.map(arg => (typeof arg === 'object' ? JSON.stringify(arg, null, 2) : arg)).join(' ');
    outputDiv.textContent += message + "\r\n";
    originalConsoleLog(FormatLog(message));
  };
}
ShowConsoleInTaskPane(document.getElementById("debugConsole"));

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

  testRecipient = addinSettings.get("testRecipient");
  if (testRecipient == null) {
    addinSettings.set("testRecipient", "");
    testRecipient = "";
    settingsUpdated = true;
  }

  // Initialise checkboxes
  addInOptions.forEach(function(addinOption) {
    settingsUpdated = settingsUpdated | InitialiseAddinOption(addinOption);
    showAddinSetting(addinOption);
    var settingCheckbox = document.getElementById(addinOption + "Checkbox");
    settingCheckbox.addEventListener("change", checkboxChanged);
  });

  if (settingsUpdated) {
    addinSettings.saveAsync(null);
    console.log("Settings written");
  } else {
    console.log("Settings read");
  }

  document.getElementById("apiUrlInput").value = apiUrl;

  var apiDelayInput = document.getElementById("apiDelayInput");
  apiDelayInput.value = apiDelay;
  apiDelayInput.onchange = UpdateApiDelay;

  var testRecipientInput = document.getElementById("testRecipient");
  testRecipientInput.value = testRecipient;
  testRecipientInput.onchange = UpdateTestRecipient;

  document.getElementById("openLink").onclick = openFolderLocationInWeb; // Add the click event for the new button
  document.getElementById("openDialog").onclick = openOfficeDialog;
  document.getElementById("applyInsightMessage").onclick = applyInsightMessage;
  document.getElementById("getMessageDetails").onclick = getMessageDetails;
  document.getElementById("sendMessage").onclick = sendMessage;
  document.getElementById("createNewAppointment").onclick = createNewAppointment;
  document.getElementById("setProps").onclick = setExtendedProperties;

  // Set up the ItemChanged event.
  if (Office.context.mailbox.item == null) {
    console.log("Item is null");
  }
  Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
  console.log("ItemChanged event handler added");
  updateTaskPaneUI(Office.context.mailbox.item);
  UpdateTestAvailability();

  initializeHTMLDragDropHandlers();
  initializeOfficeDragAndDropHandlers();
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

function UpdateTestRecipient() {
    console.log("UpdateTestRecipient called");
    var updatedTestRecipient = document.getElementById("testRecipient").value;
    if (updatedTestRecipient != testRecipient) {
        testRecipient = updatedTestRecipient;
        addinSettings.set("testRecipient", testRecipient);
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
    console.log("Couldn't locate " + checkboxName);
    return;
  }

  var checkboxLabel = document.getElementById(checkboxName + "Label");
  if (checkboxLabel == null) {
    console.log("Couldn't locate " + checkboxName + "Label");
    return;
  }

  var addinSettingValue = addinSettings.get(settingName);
  console.log(settingName + ": " + addinSettingValue);

  if ((addinSettingValue == "true" || addinSettingValue == true) && !checkbox.checked) {
    checkbox.checked = true;
    checkboxLabel.classList.add("is-checked");
  }
  else if ((addinSettingValue == "false" || addinSettingValue == false) && checkbox.checked) {
    checkbox.checked = false;
    checkbox.classList.remove("is-checked");
  }
}

function applyCheckboxSetting(settingName) {
  var checkbox = document.getElementById(settingName + "Checkbox");
  var checkboxChecked = checkbox.checked;
  console.log(settingName + " set: " + checkboxChecked);
  addinSettings.set(settingName, checkboxChecked);
}

export function checkboxChanged() {
  // Called when any checkbox is changed from the UI (we read all checkbox values and set them)
  console.log("checkboxChanged called - reading add-in settings from UI");

  addInOptions.forEach(function(addInOption) {
    applyCheckboxSetting(addInOption);
  });

  addinSettings.saveAsync(null);
}

function openURL(linkToOpen) {
  if (Office.context.ui && Office.context.ui.openBrowserWindow) {
    console.log(`Opening ${linkToOpen} using openBrowserWindow`);
    Office.context.ui.openBrowserWindow(linkToOpen);
  } else {
    console.log(`Opening ${linkToOpen} using window.open`);
    //window.open(linkToOpen, "_blank", "noopener,noreferrer");
    var newWindow = window.open("about:blank?unfiltered", "_blank");
    newWindow.location.href = linkToOpen;
  }  
}

const openFolderLocationInWeb = async () => {
  console.log("Open external link clicked");
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
  console.log(arg);
  console.log("Message received from dialog");

  if (arg.startsWith("http")) {
    // Open the external URL sent from the dialog
    openURL(arg);
  } else {
    console.log(`Unhandled message: ${arg}`);
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
        actionText: "Open TaskPane",
        actionType: Office.MailboxEnums.ActionType.ShowTaskPane,
        commandId: "msgComposeOpenPaneButton",
        contextData: "{}"
      }
    ]
  };
}

const applyInsightMessage = async () => {
  var notification = await getInsightMessage();
  if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Appointment) {
    notification.actions[0].commandId = "apptComposeOpenPaneButton";
  }

  console.log("Applying InsightMessage (from TaskPane button):");
  console.log(notification);
  Office.context.mailbox.item.notificationMessages.replaceAsync("InsightMessage", notification, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to apply InsightMessage:", asyncResult.error.message);
      return;
    }
    console.log("InsightMessage applied");
  });  
}

// This function is called when the item changes in the task pane.
function itemChanged(eventArgs) {
  // Update UI based on the new current item.
  console.log("ItemChanged event fired");
  updateTaskPaneUI(Office.context.mailbox.item);
  UpdateTestAvailability();
}

/**
 * Write the current item's subject to the TaskPane
 */
function showSubject(subject) {
    var messageSubject = document.getElementById("messageSubjectInput");
    messageSubject.value = subject;
}

function inComposeMode()
{
  if (Office.context.mailbox.item == null) {
    return false;
  }
  if (typeof Office.context.mailbox.item.subject === "string") {
    return false;
  }
  return true;
}

function UpdateTestAvailability()
{
  // Enable or disable test availablity
  if (inComposeMode()) {
    console.log("Updating available tests for Compose mode");
    document.getElementById("createNewAppointment").style.display = "none";
    document.getElementById("sendMessage").style.display = "block";
    document.getElementById("applyInsightMessage").style.display = "block";
  }
  else
  {
    console.log("Updating available tests for Read mode");
    document.getElementById("createNewAppointment").style.display = "block";
    document.getElementById("sendMessage").style.display = "none";
    document.getElementById("applyInsightMessage").style.display = "none";
  }
}

// This function updates the task pane UI based on the current item.
// All we actually do is write the subject to the console, but you could update the UI in other ways.
function updateTaskPaneUI(item) {
  if (item == null) {
    console.log("Item is null, unable to read subject");
    return;
  }

  // Because we are using the same TaskPane in both compose and read modes, we need to check which mode we are in.
  // Easy test for this is to check the type of the subject property (it will only be a string in read mode).

  // Test for read mode
  if (typeof item.subject === "string") {
    // If the subject is a string, we are in read mode.
    console.log("Item subject (read mode): " + item.subject);
    showSubject(item.subject);
    console.log("Item recipients (read mode):");
    const msgTo = item.to;
    for (let i = 0; i < msgTo.length; i++) {
      console.log(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")");
    }
    return;
  }
  
  // We are in compose mode
  item.subject.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
        return;
    }
    console.log("Item subject (compose mode): " + asyncResult.value);
    showSubject(asyncResult.value);

    if (item.itemType == Office.MailboxEnums.ItemType.Message)
    {
      Office.context.mailbox.item.to.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const msgTo = asyncResult.value;
          console.log("Item recipients (compose mode):");
          for (let i = 0; i < msgTo.length; i++) {
            console.log(msgTo[i].displayName + " (" + msgTo[i].emailAddress + ")");
          }
        } else {
          console.error(asyncResult.error);
        }
      });
    }
    else if (item.itemType == Office.MailboxEnums.ItemType.Appointment)
    {
      Office.context.mailbox.item.requiredAttendees.getAsync(function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const apptAttendees = asyncResult.value;
          console.log("Appointment attendees (compose mode):");
          for (let i = 0; i < apptAttendees.length; i++) {
            console.log(apptAttendees[i].displayName + " (" + apptAttendees[i].emailAddress + ")");
          }
        } else {
          console.error(asyncResult.error);
        }
      });            
    }
  });  
}

function getMessageDetails() {
  // Retrieve current message Id and body

  console.log("Message ID:", Office.context.mailbox.item.itemId);
  // Get the current body of the message or appointment.
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
    if (bodyResult.status === Office.AsyncResultStatus.Failed) {
      console.log(`Failed to get body: ${bodyResult.error.message}`);
      return;
    }

    console.log(bodyResult.value);
  });  
}

function sendMessage() {
  // Send the current message or appointment
  console.log("Send message selected from TaskPane");
  Office.context.mailbox.item.sendAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log("Message sent successfully");
    } else {
      console.log("Failed to send message:", result.error);
    }
  });
}

function createNewAppointment() {
  console.log("Creating new appointment");
  const start = new Date();
  const end = new Date();
  end.setHours(start.getHours() + 1);

  Office.context.mailbox.displayNewAppointmentForm({
    requiredAttendees: [],
    optionalAttendees: [],
    start: start,
    end: end,
    location: "Nowhere",
    subject: "Meeting created from add-in",
    resources: [],
    body: "Hello World!"
  });
}



function initializeHTMLDragDropHandlers() {
  // Set up drag/drop
  const dropTarget1 = document.getElementById("drop-target");
  const dropTarget2 = document.getElementById("drop-target-2");

  dropTarget1.addEventListener("dragenter", (event) => {
    event.preventDefault();
    console.log("Target 1: dragenter");
    dropTarget1.style.backgroundColor = "lightblue";
  });
  dropTarget1.addEventListener("dragleave", (event) => {
    event.preventDefault();
    console.log("Target 1: dragleave");
    dropTarget1.style.backgroundColor = "lightgreen";
  });
  dropTarget1.addEventListener("drop", (event) => {
    event.preventDefault();
    console.log("Target 1: drop");
    dropTarget1.style.backgroundColor = "lightgreen";
    const data = event.dataTransfer.getData("text/plain");
    console.log("Dropped data: " + data);
  });  

  dropTarget2.addEventListener("dragover", (event) => {
    event.preventDefault();
    console.log("Target 2: dragenter");
    dropTarget2.style.backgroundColor = "lightcoral";
  });
  dropTarget2.addEventListener("dragleave", (event) => {
    event.preventDefault();
    console.log("Target 2: dragleave");
    dropTarget2.style.backgroundColor = "lightsalmon";
  });
  dropTarget2.addEventListener("drop", (event) => {
    event.preventDefault();
    console.log("Target 2: drop");
    dropTarget2.style.backgroundColor = "lightsalmon";
    const data = event.dataTransfer.getData("text/plain");
    console.log("Dropped data: " + data);
  });  
}

let parentEntryWidth = 30000;

function initializeOfficeDragAndDropHandlers() {
       // Handle the DragAndDropEvent event.
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.DragAndDropEvent,
      (event) => {
        const eventData = event.dragAndDropEventData;

        if(eventData.type == "dragover"){
          const {possibleX, possibleY} = reScaleCoordinates(eventData.pageX, eventData.pageY);
          const elementUnderCursor = document.elementFromPoint(possibleX, possibleY);
          
          if(elementUnderCursor){
            elementUnderCursor.dispatchEvent(new DragEvent('dragover', {
                    bubbles: true,
                    cancelable: true,
                    clientX: possibleX,
                    clientY: possibleY,
              }));
          }  
        }

        // Get the file name and the contents of the items dropped into the task pane.
        if (eventData.type == "drop") {
          console.log(eventData);
          //console.log("pageX: " + eventData.pageX + ", pageY: " + eventData.pageY);
          const files = eventData.dataTransfer.files;
          files.forEach((file) => {
            const content = file.fileContent;
            const name = file.name;

            // Add operations to process the item here, such as uploading the file to a CRM system.
            console.log(`File name: ${name}, File content: ${content}`);
          });
        }
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to add Office.EventType.DragAndDropEvent handler:", asyncResult.error.message);
          return;
        }

        console.log("Event handler added successfully.");
      }
    );
}

function reScaleCoordinates(pageX, pageY){
  // Issue: This will only work if user is draging from left to right.
  parentEntryWidth = Math.min(pageX, parentEntryWidth);
  const possibleX = pageX - parentEntryWidth;

  //Issue: The value 113 is hardcoded for now. Need to find a way to get this value dynamically.
  const possibleY = pageY - 113;

  //var iframeRect = window.frameElement.getBoundingClientRect();
  //console.log(`iFrame Left: ${iframeRect.left}, Top: ${iframeRect.top}`);

  const width = document.documentElement.clientWidth;
  const height = document.documentElement.clientHeight;
  //console.log(`Frame Width: ${width}, Height: ${height}`);  

  //var fOffset = computeFrameOffset(window);
  //console.log(`Frame Offset Left: ${fOffset.left}, Top: ${fOffset.top}`);
  //console.log(`window.top: ${window.top}`);
  //console.log(`window.parent: ${window.parent}`);

  return { possibleX, possibleY };
}

function gP(e){var left=0;var top=0; while (e.offsetParent){ left+=e.offsetLeft-e.scrollLeft;top+=e.offsetTop-e.scrollTop;e=e.offsetParent;}return {x:left, y:top};}
window.getPos = gP;

/**
 * Calculate the offset of the given iframe relative to the top window.
 * - Walks up the iframe chain, checking the offset of each one till it reaches top
 * - Only works with friendly iframes. https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy#Cross-origin_script_API_access 
 * - Takes into account scrolling, but comes up with a result relative to 
 *   top iframe, regardless of being visibile withing intervening frames.
 * 
 * @param window win    the iframe we're interested in (e.g. window)
 * @param object dims   an object containing the offset so far:
 *                          { left: [x], top: [y] }
 *                          (optional - initializes with 0,0 if undefined) 
 * @return dims object above
 */
var computeFrameOffset = function(win, dims) {
    // initialize our result variable
    if (typeof dims === 'undefined') {
        var dims = { top: 0, left: 0 };
    }

    // find our <iframe> tag within our parent window
    var frames = win.parent.document.getElementsByTagName('iframe');
    var frame;
    var found = false;

    for (var i=0, len=frames.length; i<len; i++) {
        frame = frames[i];
        if (frame.contentWindow == win) {
            found = true;
            break;
        }
    }

    // add the offset & recur up the frame chain
    if (found) {
        var rect = frame.getBoundingClientRect();
        dims.left += rect.left;
        dims.top += rect.top;
        if (win !== top) {
            computeFrameOffset(win.parent, dims);
        }
    }
    return dims;
};

function SetDefaultMessageProperties(subject) {
  // Check if the message has a recipient.  If not, add 1@demonmaths.co.uk as the recipient
  Office.context.mailbox.item.to.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = asyncResult.value;
      if (recipients.length === 0 && testRecipient.length > 0) {
        // No recipients found, add default recipient
        Office.context.mailbox.item.to.addAsync([testRecipient], (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Default recipient added.");
          } else {
            console.error("Failed to add default recipient:", asyncResult.error.message);
          }
        });
      }
    } else {
      console.error("Failed to get recipients:", asyncResult.error.message);
    }
  });

  // Check if the message has a subject.  If not, set it to the given value with the current date and time
  if (!subject || subject.length === 0) {
    subject = "Test";
  }
  Office.context.mailbox.item.subject.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const subject = asyncResult.value;
      if (!subject) {
        const currentDateTime = new Date().toISOString();
        Office.context.mailbox.item.subject.setAsync(`${subject} - ${currentDateTime}`, (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Subject set successfully.");
          } else {
            console.error("Failed to set subject:", asyncResult.error.message);
          }
        });
      }
    } else {
      console.error("Failed to get subject:", asyncResult.error.message);
    }
  });

}

async function setExtendedProperties() {
  SetDefaultMessageProperties("Testing Extended Properties");

  // Set extended property named daves.tips with a value set to current date and time
  console.log("Attempting to set extended property");
  const currentDateTime = new Date().toISOString();
  // Load extended properties
  await sleep(2000);
  Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      const customProps = asyncResult.value;      
      customProps.set("daves.tips", currentDateTime);
      console.log(`Set extended property daves.tips to ${currentDateTime}`);
      customProps.saveAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Extended properties saved successfully.");
          // Prepend success to message body
          Office.context.mailbox.item.body.prependAsync(`Successfully set extended property daves.tips to ${currentDateTime}\n\n`, (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("Message body prepended successfully.");
              Office.context.mailbox.item.saveAsync();
            } else {
              console.error("Failed to prepend message body:", asyncResult.error.message);
            }
          });
        } else {
          console.error("Failed to save extended properties:", asyncResult.error.message);
        }
      });
    } else {
      console.error("Failed to load custom properties:", asyncResult.error.message);
    }
  });
}

/**
 * Sleep function for JavaScript
 * @param {number} ms - Time to sleep in milliseconds
 * @returns {Promise<void>}
 */
function sleep(ms) {
    return new Promise((resolve, reject) => {
        // Validate input
        if (typeof ms !== 'number' || ms < 0 || !Number.isFinite(ms)) {
            reject(new Error("Invalid delay time. Must be a non-negative number."));
            return;
        }
        setTimeout(resolve, ms);
    });
}