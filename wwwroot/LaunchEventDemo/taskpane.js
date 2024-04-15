/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * The name of the add-in.
 * @type {string}
 */
const AddinName = "LaunchEventDemo";

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
    }
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
}

/**
 * Updates the delay for the API call.
 */
function UpdateApiDelay() {
    var apiDelay = document.getElementById("apiDelayInput").value;
    if (apiDelay != apiDelayInSeconds) {
        apiDelayInSeconds = Number(apiDelay);
        addinSettings.set("apiDelay", apiDelayInSeconds);
        addinSettings.saveAsync(null);
    }
}