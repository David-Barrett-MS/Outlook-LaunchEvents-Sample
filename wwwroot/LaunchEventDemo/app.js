/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

var statusInfo = "";
var fullLogEventAPIUrl = ""; // The API URL including any additional parameters
var baseLogEventAPIUrl = ""; // The API URL
var addinSettings;
const AddinName = "LaunchEventDemo";

//Office.onReady();
Office.initialize = function () {
    // This function seems not to be called during OnMessageSend LaunchEvent, so any initialisation here won't work in that scenario
    ReadAddinSettings();
}

/**
 * Reads the add-in settings and updates the fullLogEventAPIUrl variable accordingly.
 */
function ReadAddinSettings() {
    if (!addinSettings) {
        addinSettings = Office.context.roamingSettings;
    }

    if (baseLogEventAPIUrl == "") {
        baseLogEventAPIUrl = addinSettings.get("apiUrl");
        if (baseLogEventAPIUrl) {
            fullLogEventAPIUrl = baseLogEventAPIUrl;
            LogToConsole("API URL read: " + fullLogEventAPIUrl);
        }    
        
        let apiDelay = 0;
        apiDelay = addinSettings.get("apiDelay");
        LogToConsole("API delay: " + apiDelay);
        if (apiDelay > 0) {
            fullLogEventAPIUrl = fullLogEventAPIUrl + "?DelayInSeconds=" + apiDelay
            LogToConsole("API URL adjusted: " + fullLogEventAPIUrl);
        }
        LogToConsole("Finished reading add-in settings");
    }
}

/**
 * Set notification on MailItem (overwrites any previous notification)
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
 * Append the given status to the notification for the MailItem
 * @param {Message to be added to the status} message 
 * @returns 
 */
async function SetStatus(message) {
    if (statusInfo != "") {
        statusInfo = statusInfo + " | ";    
    }
    statusInfo = statusInfo + message;
    console.log(message);
    return SetNotification(statusInfo);
}

function LogToConsole(data) {
    console.log(AddinName + ": " + data);
}

/**
 * Logs an event and sends it to the server. allowEvent is always set to true, and we don't wait for server response.
 * @param {string} eventData - The data to be sent to the server (e.g. event name).
 * @param {object} event - The Outlook event object (to be marked completed when done).
 * @returns {Promise<void>} - A promise that resolves when the request has been sent to the server.
 */
async function logEvent(eventData, event) {
    ReadAddinSettings();
    if (baseLogEventAPIUrl != "") {
        LogToConsole("POST " + baseLogEventAPIUrl);
        eventData = AddinName + ": " + eventData;
        var xhr = new XMLHttpRequest();
        xhr.timeout = 300000;
        xhr.open("POST", baseLogEventAPIUrl, true);
        xhr.setRequestHeader("Content-Type", "text/plain; charset=UTF-8"); 
        xhr.send(eventData);    
    }

    if (event != null) {
        event.completed({ allowEvent: true });
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
    if (fullLogEventAPIUrl != "") {
        LogToConsole("POST " + fullLogEventAPIUrl);
        eventData = AddinName + ": " + eventData;
        var xhr = new XMLHttpRequest();
        xhr.onreadystatechange = function () {
            if (this.readyState == 4) {
                if (this.status == 200 ) {
                    event.completed({ allowEvent: true });
                }
                else {
                    event.completed({ allowEvent: false, errorMessage:"Failed to contact API" });
                }
            }
        }     
        xhr.timeout = 300000; // The maximum time that Outlook allows for an event based add-in to complete the event
        xhr.open("POST", fullLogEventAPIUrl, true);
        xhr.setRequestHeader("Content-Type", "text/plain; charset=UTF-8"); 
        xhr.send(eventData);
    } else {
        event.completed({ allowEvent: false, errorMessage:"API URL not set - open TaskPane to configure" });
    }
}

// <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="SoftBlock"/>
function onMessageSendHandler(event) {
    logEvent2("OnMessageSend", event);
}

// <LaunchEvent Type="OnNewMessageCompose" FunctionName="OnNewMessageComposeHandler"/>
function OnNewMessageComposeHandler(event) {
    logEvent("OnNewMessageCompose", event);
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

// <LaunchEvent Type="OnAppointmentSend" FunctionName="OnAppointmentSendHandler" SendMode="Block"/>
function OnAppointmentSendHandler(event) {
    logEvent2("OnAppointmentSend", event);
}

// <LaunchEvent Type="OnMessageCompose" FunctionName="OnMessageComposeHandler"/>
function OnMessageComposeHandler(event) {
    logEvent("OnMessageCompose", event);
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


if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    // Associate the events with their respective handlers
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
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