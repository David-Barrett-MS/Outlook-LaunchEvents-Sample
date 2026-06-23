/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


console.log("LaunchEvent Add-in Loaded");

var statusInfo = "";
var fullLogEventAPIUrl = ""; // The API URL including any additional parameters
var baseLogEventAPIUrl = ""; // The API URL
var addinSettings;
var clientDelayInSeconds = 0;
const AddinName = "LaunchEventDemo";

//Office.onReady();
Office.initialize = function () {
    // This function is not called during OnMessageSend LaunchEvent in Outlook Desktop, so any initialisation here won't work in that scenario
    ReadAddinSettings();
}

async function sleep(ms) {
    return new Promise((resolve, reject) => {
        if (typeof ms !== 'number' || ms < 0 || !Number.isFinite(ms)) {
            return reject(new Error("Invalid delay: must be a non-negative number."));
        }
        setTimeout(resolve, ms);
    });
}

// Try and trigger timing related error - we add a delay here
var startWaitTime = Date.now();
//for (let i = 0; i < 10000000000; i++) {
//    if (i % 100000000 == 0) {
//        console.log("Waiting...");
//    }
//}
//(async () => { await sleep(10000); })();

console.log("Done waiting, waited for " + (Date.now() - startWaitTime)/1000 + " seconds");

/**
 * Reads the add-in settings and updates the fullLogEventAPIUrl variable accordingly.
 */
function ReadAddinSettings() {
    if (!addinSettings) {
        addinSettings = Office.context.roamingSettings;
    } else {
        return; // We only need to read the settings once
    }

    if (baseLogEventAPIUrl == "") {
        baseLogEventAPIUrl = addinSettings.get("apiUrl");
        if (baseLogEventAPIUrl) {
            fullLogEventAPIUrl = baseLogEventAPIUrl;
            console.log(FormatLog("API URL read: " + fullLogEventAPIUrl));
        }    
        
        let apiDelay = 0;
        apiDelay = addinSettings.get("apiDelay");
        console.log(FormatLog("API delay: " + apiDelay));
        if (apiDelay > 0) {
            fullLogEventAPIUrl = fullLogEventAPIUrl + "?DelayInSeconds=" + apiDelay
            console.log(FormatLog("API URL adjusted: " + fullLogEventAPIUrl));
        }

        clientDelayInSeconds = addinSettings.get("clientDelay");
        console.log(FormatLog("Client delay: " + clientDelayInSeconds));

        console.log(FormatLog("Finished reading add-in settings"));
    }
}

/**
 * Set notification on current item (overwrites any previous notification)
 * @param {Notification message to be set} message 
 */
async function SetNotification(message) {
    let infoMessage =
    {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "icon2",
      persistent: true
    };    
    Office.context.mailbox.item.notificationMessages.replaceAsync(AddinName + "Notification", infoMessage);
}

/**
 * Append the given status to the notification for the current item
 * @param {Message to be added to the status} message 
 * @returns 
 */
async function SetStatus(message) {
    if (statusInfo != "") {
        statusInfo = statusInfo + " | ";    
    }
    statusInfo = statusInfo + message;
    console.log(FormatLog("Adding to notification: " + message));
    return SetNotification(statusInfo);
}

function FormatLog(data) {
    // Return log with add-in name and current time prepended
    let now = new Date();
    let currentTime = now.toLocaleTimeString('en-US', { hour12: false });
    return AddinName + " " + currentTime + ": " + data;
}

/**
 * Logs an event and sends it to the server. allowEvent is always set to true, and we don't wait for server response.
 * @param {string} eventData - The data to be sent to the server (e.g. event name).
 * @param {object} event - The Outlook event object (to be marked completed when done).
 * @returns {Promise<void>} - A promise that resolves when the request has been sent to the server.
 */
async function logEvent(eventData, event) {
    ReadAddinSettings();
    console.log(FormatLog(eventData + " received"));
    if (addinSettings.get("showEventsOnMessage") == "true" || addinSettings.get("showEventsOnMessage") == true) {
        SetStatus(eventData);
    }

    if (baseLogEventAPIUrl != "") {
        console.log(FormatLog("POST " + baseLogEventAPIUrl));
        
        if (addinSettings.get("sendClientInfo") == "true" || addinSettings.get("sendClientInfo") == true) {
            eventData = Office.context.mailbox.userProfile.displayName + ": " + eventData;
        } else {
            eventData = AddinName + ": " + eventData;
        }
        var xhr = new XMLHttpRequest();
        xhr.timeout = 300000;
        xhr.open("POST", baseLogEventAPIUrl, true);
        xhr.setRequestHeader("Content-Type", "text/plain; charset=UTF-8");
        xhr.send(eventData);
    }

    if (event != null) {
        logEventCompleted(eventData, event);
    }
}

/**
 * Logs an event and sends it to the server.  If an error occurs, allowEvent is set to false on event completion.
 * @param {string} eventData - The data to be sent to the server (e.g. event name).
 * @param {object} event - The Outlook event object (to be marked completed when done).
 * @returns {Promise<void>} - A promise that resolves when the event is logged.
 */
async function logEvent2(eventData, event) {
    ReadAddinSettings();
    console.log(FormatLog(eventData + " received"));
    if (fullLogEventAPIUrl != "") {
        console.log(FormatLog("POST " + fullLogEventAPIUrl));
        if (addinSettings.get("sendClientInfo") == "true" || addinSettings.get("sendClientInfo") == true) {
            eventData = Office.context.mailbox.userProfile.displayName + ": " + eventData;
        } else {
            eventData = AddinName + ": " + eventData;
        }
        
        // Check for sessionData and append if available (for send events)
        if (eventData.includes("Send") && Office.context.mailbox.item.sessionData) {
            Office.context.mailbox.item.sessionData.getAllAsync((asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    const sessionData = asyncResult.value;
                    if (sessionData && Object.keys(sessionData).length > 0) {
                        console.log(FormatLog("SessionData found:"));
                        for (const key in sessionData) {
                            console.log(FormatLog(`  ${key} = ${sessionData[key]}`));
                            eventData += ` | SessionData: ${key}=${sessionData[key]}`;
                        }
                    }
                }
            });
        }
        sendLogEventRequest(eventData, event);
    } else {
        if (event != null) {
            logEventCompleted(eventData, event, "API URL not set - open TaskPane to configure");
        }
    }
}

function sendLogEventRequest(eventData, event) {
    var xhr = new XMLHttpRequest();
    xhr.onreadystatechange = function () {
        if (this.readyState == 4) {
            if (this.status == 200 || (addinSettings.get("blockOnAPIFail") != true && addinSettings.get("blockOnAPIFail") != "true") ) {
                logEventCompleted(eventData, event);
            }
            else {
                logEventCompleted(eventData, event, "Failed to contact API");
            }
        }
    }
    xhr.timeout = 300000; // The maximum time that Outlook allows for an event based add-in to complete the event
    xhr.open("POST", fullLogEventAPIUrl, true);
    xhr.setRequestHeader("Content-Type", "text/plain; charset=UTF-8"); 
    xhr.send(eventData);
}

function logEventCompleted(eventData, event, errorMessage) {
    if (errorMessage) {
        console.log(FormatLog(eventData + ": Sending event.completed with error: " + errorMessage));
        if (event != null) { event.completed({ allowEvent: false, errorMessage: errorMessage }); }
    } else {
        console.log(FormatLog(eventData + ": Sending event.completed"));
        if (event != null) { event.completed({ allowEvent: true }); }
    }
}

// <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock"/>
function onMessageSendHandler(event) {
    //applyInsightMessage(null);
    ReadAddinSettings();

    // Read showCustomSmartAlertDialog setting dynamically from add-in settings
    var showSmartAlert = addinSettings.get("showCustomSmartAlertDialog");
    if (showSmartAlert == "true" || showSmartAlert == true) {
        // Check if send requirements have been met
        Office.context.mailbox.item.loadCustomPropertiesAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                const customProps = asyncResult.value;
                const sendRequirementsMet = customProps.get("sendRequirementsMet");
                
                if (sendRequirementsMet === "true" || sendRequirementsMet === true) {
                    console.log(FormatLog("Send requirements met, allowing send"));
                    // Requirements met, proceed with normal send flow
                    if (clientDelayInSeconds > 0) {
                        setTimeout(() => {
                            logEvent2("OnMessageSend", event);
                        }, clientDelayInSeconds * 1000);
                    } else {
                        logEvent2("OnMessageSend", event).then(() => { logEventCompleted("OnMessageSend", event);});
                    }
                } else {
                    console.log(FormatLog("Send requirements not met, blocking send"));
                    // Requirements not met, show smart alert
                    event.completed({
                        allowEvent: false,
                        errorMessage: "Please ensure send requirements are met in the TaskPane.",
                        errorMessageMarkdown: "Please ensure send requirements are met in the TaskPane.",
                        cancelLabel: "Open TaskPane",
                        commandId: "msgComposeOpenPaneButton"
                    });
                }
            } else {
                console.log(FormatLog("Failed to load custom properties, blocking send"));
                // Failed to load properties, block send to be safe
                event.completed({
                    allowEvent: false,
                    errorMessage: "Failed to validate send requirements.",
                    errorMessageMarkdown: "Failed to validate send requirements.",
                    cancelLabel: "Open TaskPane",
                    commandId: "msgComposeOpenPaneButton"
                });
            }
        });
        return;
    }


    setTimeout(() => {
        console.log("Timer fired AFTER completed");
    }, 15000);


    if (clientDelayInSeconds > 0) {
        setTimeout(() => {
            logEvent2("OnMessageSend", event);
        }, clientDelayInSeconds * 1000);
    } else {
        logEvent2("OnMessageSend", event).then(() => { logEventCompleted("OnMessageSend", event);});
    }
}


// <LaunchEvent Type="OnAppointmentSend" FunctionName="OnAppointmentSendHandler" SendMode="Block"/>
function OnAppointmentSendHandler(event) {
    //applyInsightMessage(null);
    logEvent2("OnAppointmentSend", event);
}

// <LaunchEvent Type="OnNewMessageCompose" FunctionName="OnNewMessageComposeHandler"/>
function OnNewMessageComposeHandler(event) {
    logEvent("OnNewMessageCompose", null).then(() => {
        logEventCompleted("OnNewMessageCompose",event);
    })
}

//<LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="OnNewAppointmentOrganizerHandler"/>
function OnNewAppointmentOrganizerHandler(event) {
    logEvent("OnNewAppointmentOrganizer", event);
}

// <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="OnMessageAttachmentsChangedHandler"/>
function OnMessageAttachmentsChangedHandler(event) {
    logEvent("OnMessageAttachmentsChanged", event);
}

// <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="OnAppointmentAttachmentsChangedHandler"/>
function OnAppointmentAttachmentsChangedHandler(event) {
    logEvent("OnAppointmentAttachmentsChanged", event);
}

//<LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="OnMessageRecipientsChangedHandler"/>
function OnMessageRecipientsChangedHandler(event) {
    logEvent("OnMessageRecipientsChanged", event);
}

// <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="OnAppointmentAttendeesChangedHandler"/>
function OnAppointmentAttendeesChangedHandler(event) {
    logEvent("OnAppointmentAttendeesChanged", event);
}

// <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="OnAppointmentTimeChangedHandler"/>
function OnAppointmentTimeChangedHandler(event) {
    logEvent("OnAppointmentTimeChanged", event);
}

// <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="OnAppointmentRecurrenceChangedHandler"/>
function OnAppointmentRecurrenceChangedHandler(event) {
    logEvent("OnAppointmentRecurrenceChanged", event);
}

// <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="OnInfoBarDismissClickedHandler"/>
function OnInfoBarDismissClickedHandler(event) {
    logEvent("OnInfoBarDismissClicked", event);
}


// <LaunchEvent Type="OnMessageCompose" FunctionName="OnMessageComposeHandler"/>
function OnMessageComposeHandler(event) {
    logEvent("OnMessageCompose", null).then(() => {
        //logEventCompleted("OnMessageCompose",event);
    });
}

// <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="OnAppointmentOrganizerHandler"/>
function OnAppointmentOrganizerHandler(event) {
    logEvent("OnAppointmentOrganizer", event);
}

// <LaunchEvent Type="OnMessageFromChanged" FunctionName="OnMessageFromChangedHandler"/>
function OnMessageFromChangedHandler(event) {
    logEvent("OnMessageFromChanged", event);
}

// <LaunchEvent Type="OnAppointmentFromChanged" FunctionName="OnAppointmentFromChangedHandler"/>
function OnAppointmentFromChangedHandler(event) {
    logEvent("OnAppointmentFromChanged", event);
}

// <LaunchEvent Type="OnSensitivityLabelChanged" FunctionName="OnSensitivityLabelChangedHandler"/>
function OnSensitivityLabelChangedHandler(event) {
    logEvent("OnSensitivityLabelChanged", event);
}

// <LaunchEvent Type="OnMessageReadWithCustomAttachment" FunctionName="OnMessageReadWithCustomAttachmentHandler"/>
function OnMessageReadWithCustomAttachmentHandler(event) {
    logEvent("OnMessageReadWithCustomAttachment", event);
}

// <LaunchEvent Type="OnMessageReadWithCustomHeader" FunctionName="OnMessageReadWithCustomHeaderHandler"/>
function OnMessageReadWithCustomHeaderHandler(event) {
    logEvent("OnMessageReadWithCustomHeader", event);
}


if (Office.context !== undefined && (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) ) {
    // Associate the events with their respective handlers
    Office.actions.associate("OnMessageSendHandler", onMessageSendHandler);
    Office.actions.associate("OnNewMessageComposeHandler", OnNewMessageComposeHandler);
    Office.actions.associate("OnNewAppointmentOrganizerHandler", OnNewAppointmentOrganizerHandler);
    Office.actions.associate("OnMessageAttachmentsChangedHandler", OnMessageAttachmentsChangedHandler);
    Office.actions.associate("OnAppointmentAttachmentsChangedHandler", OnAppointmentAttachmentsChangedHandler);
    Office.actions.associate("OnMessageRecipientsChangedHandler", OnMessageRecipientsChangedHandler);
    Office.actions.associate("OnAppointmentAttendeesChangedHandler", OnAppointmentAttendeesChangedHandler);
    Office.actions.associate("OnAppointmentTimeChangedHandler", OnAppointmentTimeChangedHandler);
    Office.actions.associate("OnAppointmentRecurrenceChangedHandler", OnAppointmentRecurrenceChangedHandler);
    Office.actions.associate("OnInfoBarDismissClickedHandler", OnInfoBarDismissClickedHandler);
    Office.actions.associate("OnAppointmentSendHandler", OnAppointmentSendHandler);
    Office.actions.associate("OnMessageComposeHandler", OnMessageComposeHandler);
    Office.actions.associate("OnAppointmentOrganizerHandler", OnAppointmentOrganizerHandler);
    Office.actions.associate("OnMessageFromChangedHandler", OnMessageFromChangedHandler);
    Office.actions.associate("OnAppointmentFromChangedHandler", OnAppointmentFromChangedHandler);
    Office.actions.associate("OnSensitivityLabelChangedHandler", OnSensitivityLabelChangedHandler);
    // Office.actions.associate("OnMessageReadWithCustomAttachmentHandler", OnMessageReadWithCustomAttachmentHandler);
    // Office.actions.associate("OnMessageReadWithCustomHeaderHandler", OnMessageReadWithCustomHeaderHandler);
}


