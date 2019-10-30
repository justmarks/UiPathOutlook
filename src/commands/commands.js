/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
 
Office.onReady(info => {
  // Initialize the robot
  // UiPathRobot.init();

  // If needed, Office.js is ready to be called

   //Find the right process
   processId = GetProcessByName(processName);
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  DisplayMessage("Starting...");

  var item = Office.context.mailbox.item;
  var listOfAttachments = [];
  var options = {asyncContext: {currentItem: item}};
  var attachments = item.attachments;
  if (attachments.length > 0) {
    attachments.forEach(async function (attachment) {
      DisplayMessage("Downloading invoice " + attachment.name);
      item.getAttachmentContentAsync(attachment.id, handleAttachmentsCallback);
    });
  }

  DisplayMessage("Processing complete")

  // Be sure to indicate when the add-in command function is complete
  event.completed();
} 

async function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      /*fs.writeFile(fileName, result.value, () => {
        DisplayMessage(result.name + " stored");
      });*/
      DisplayMessage("Invoice downloaded.  Invoking robot to analyze invoice...");

      // Run Process, passing argument in to folder
      UiPathRobot.runProcess(processId, arguments, (status) => {
        DisplayMessage(status);
      }).then(
        result => { DisplayMessage("Invoice processed!") },
        err => { DisplayMessage("Error:" + err); }
      );
      break;
  };
};

function getGlobal() { 
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window : 
    (typeof global !== "undefined") ? global :
    undefined;
}

function DisplayMessage(messageText) {
  console.log(messageText);

  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage, 
    message: messageText,
    icon: "Icon.80x80",
    persistent: true
  }
  
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  console.log(messageText);
}

function GetProcessByName(processName) {
  UiPathRobot.getRobotProcesses().then((processes) => {
    for (var i=0; i<processes.length; i++) {
      console.log(processes[i].name.toLowerCase());
      if (processes[i].name.toLowerCase().startsWith(processName.toLowerCase())) {
        console.log("found");
        return processes[i].id;
      }
    }
  });
}
  

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;

var localFolder = "Invoices";
var processId = "";

const processName = "relocation_invoice_processing";