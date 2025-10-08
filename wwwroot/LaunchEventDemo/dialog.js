import { externalLink } from "./common/constants.js";

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("openLinkViaParent").onclick = openExternalURLViaParent; // Add the click event for the new button
  document.getElementById("openLinkDirect").onclick = openExternalURL; // Add the click event for the new button
});

const openExternalURLViaParent = async () => {
  console.log(`Sending external URL to parent: ${externalLink}`);
  
  // Send the external URL back to the parent taskpane
  Office.context.ui.messageParent(externalLink);
};

const openExternalURL = async () => {
  if (Office.context.ui.openBrowserWindow === undefined) {
    console.log(`Opening external URL using window.open: ${externalLink}`);
    var newWindow = window.open("about:blank?unfiltered", "_blank");
    newWindow.location.href = externalLink;
    //window.open(externalLink, "_self");
  } else {
    console.log(`Opening external URL using Office openBrowserWindow: ${externalLink}`);
    Office.context.ui.openBrowserWindow(externalLink);
  }
};
