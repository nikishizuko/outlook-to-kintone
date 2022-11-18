/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

let item;
let settings;
const sentMessageElement = document.getElementById("sent-message");
const saveMessageElement = document.getElementById("save-message");

Office.initialize = () => {};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("settings-close-icon").style.display = "none";
    document.getElementById("settings-section").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("save-settings").onclick = saveSettings;
    document.getElementById("send-data").onclick = sendData;
    document.getElementById("settings-open-icon").onclick = openSettings;
    document.getElementById("settings-close-icon").onclick = closeSettings;
    item = Office.context.mailbox.item;
    settings = Office.context.roamingSettings;

    // Set current settings as existing saved webhook if exists
    document.getElementById("webhook-url").value = settings.get("webhookUrl") ? settings.get("webhookUrl") : "";
  }
});

// Handler when "Save" button is clicked in settings
function saveSettings() {
  const webhookInput = document.getElementById("webhook-url").value;
  settings.set("webhookUrl", webhookInput);
  settings.saveAsync(saveSettingsCallback);

  // Display saved message after saving
  saveMessageElement.innerHTML = "<b>Saved!</b> <br/>";
  setTimeout(() => {
    saveMessageElement.innerHTML = "";
  }, 3000);
}

// Callback function for the setting save
function saveSettingsCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // Display error if error occurred
    saveMessageElement.innerHTML = `<span style="font-weight: bold; color: red">ERROR: ${asyncResult.error.message}</span>`;
    setTimeout(() => {
      saveMessageElement.innerHTML = "";
    }, 3000);
  }
}

// Handler when gear icon is clicked
function openSettings() {
  document.getElementById("settings-section").style.display = "block";
  document.getElementById("settings-close-icon").style.display = "block";
  document.getElementById("settings-open-icon").style.display = "none";
  document.getElementById("kintone-send-section").style.display = "none";
}

// Handler when X icon is clicked
function closeSettings() {
  document.getElementById("settings-section").style.display = "none";
  document.getElementById("settings-open-icon").style.display = "block";
  document.getElementById("settings-close-icon").style.display = "none";
  document.getElementById("kintone-send-section").style.display = "block";
}

// Handler when "Send" button is clicked
async function sendData() {
  // Initialize body structure to send to Integromat
  let sendBody = {};

  // Add email(body) text to the send body
  sendBody.bodyText = document.getElementById("body-text").value;
  // Add retrieved recipients to the send body
  sendBody.recipients = await getAllRecipients();
  // Add retreived subject to the send body
  sendBody.subject = await getSubject();

  // If no webhook is set, show error to set webhook
  if (settings.get("webhookUrl") === "") {
    // Display sent message after sending
    sentMessageElement.innerHTML =
      '<span style="font-weight: bold; color: red">ERROR: Webhook URL not entered.</span><br/>';
    setTimeout(() => {
      sentMessageElement.innerHTML = "";
    }, 3000);
    return;
  }

  // XHR request to send to Integromat's webhook
  const xhr = new XMLHttpRequest();
  xhr.open("POST", settings.settingsData.webhookUrl);
  xhr.send(JSON.stringify(sendBody));

  // Display sent message after sending
  sentMessageElement.innerHTML = "<b>Sent!</b> <br/>";
  setTimeout(() => {
    sentMessageElement.innerHTML = "";
  }, 3000);
}

// Retrieves recipients for all recipient types of the email and returns object of arrays
async function getAllRecipients() {
  let recipients = {};

  const to = await getRecipients(item.to);
  const cc = await getRecipients(item.cc);
  let bcc;

  if (item.bcc) {
    bcc = await getRecipients(item.bcc);
  } else {
    bcc = [];
  }

  recipients.to = to;
  recipients.cc = cc;
  recipients.bcc = bcc;

  return recipients;
}

// Retrieves recipient data for the specified recipient type and returns array of those recipients
function getRecipients(recipientType) {
  return new Promise(function (resolve, reject) {
    try {
      recipientType.getAsync((asyncResult) => {
        let recipientsContainer = [];
        asyncResult.value.forEach((email) => {
          recipientsContainer.push(email.emailAddress);
        });
        resolve(recipientsContainer);
      });
    } catch (error) {
      console.log(error);
      reject(error);
    }
  });
}

function getSubject() {
  return new Promise(function (resolve, reject) {
    try {
      item.subject.getAsync((asyncResult) => {
        resolve(asyncResult.value);
      });
    } catch (error) {
      console.log(error);
      reject(error);
    }
  });
}
