import { externalLink } from "./common/constants.js";

/* global document, Office, console */

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("openLinkViaParent").onclick = openExternalURLViaParent; // Add the click event for the new button
  document.getElementById("openLinkDirect").onclick = openExternalURL; // Add the click event for the new button
});

async function run() {
  console.log(`Hello from the dialog box`);
}

const openExternalURLViaParent = async () => {
  console.log(`Sending external URL to parent: ${externalLink}`);
  
  // Send the external URL back to the parent taskpane
  Office.context.ui.messageParent(externalLink);
};


const openExternalURL = async () => {
  console.log(`Opening external URL directly: ${externalLink}`);
  
  window.open(externalLink, "_blank", "noopener,noreferrer");
};
