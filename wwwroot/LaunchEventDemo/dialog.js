import { externalLink } from "./common/constants.js";

/* global document, Office, console */

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("openLink").onclick = openFolderLocationInWeb; // Add the click event for the new button
});

async function run() {
  console.log(`Hello from the dialog box`);
}

const openFolderLocationInWeb = async () => {
  console.log(`Sending external URL to parent: ${externalLink}`);
  
  // Send the external URL back to the parent taskpane
  Office.context.ui.messageParent(externalLink);
};
