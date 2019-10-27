/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
 
Office.onReady(info => {
  // Initialize the robot
  UiPathRobot.init(1000);

  // If needed, Office.js is ready to be called
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
      DisplayMessage("Saving " + attachment.name);
      item.getAttachmentContentAsync(attachment.id, handleAttachmentsCallback);
    });
  }

  DisplayMessage("Processing complete")

  // Be sure to indicate when the add-in command function is complete
  event.completed();
} 

async function handleAttachmentsCallback(result) {
  DisplayMessage("start callback");
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      /*DisplayMessage("here " + result.value);

      fs.writeFile(fileName, result.value, () => {
        DisplayMessage(result.name + " stored");
      });
      DisplayMessage("done");*/
      DisplayMessage("Invoice downloaded for analysis.");

      DisplayMessage("Invoking robot...");
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

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;

var localFolder = "Invoices";

const processName = "Relocation_invoice_processing_DemoEnvironment";
const processId = "41cc2677-ebc5-46b9-bbc7-1f1ebb16d831";