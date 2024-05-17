/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onMessageSendHandler(event) {
  Office.context.mailbox.item.body.getAsync(
    "text",
    { asyncContext: event },
    getBodyCallback
  );
}

function getBodyCallback(asyncResult){
  const event = asyncResult.asyncContext;
  let body = "";
  if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
    body = asyncResult.value;
  } else {
    const message = "Failed to get body text";
    console.error(message);
    event.completed({ allowEvent: false, errorMessage: message });
    return;
  }

  const matches = hasMatches(body);
  if (matches) {
    Office.context.mailbox.item.getAttachmentsAsync(
      { asyncContext: event },
      getAttachmentsCallback);
  } else {
    event.completed({ allowEvent: true });
  }
}

function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }

  const arrayOfTerms = ["send", "picture", "document", "attachment"];
  for (let index = 0; index < arrayOfTerms.length; index++) {
    const term = arrayOfTerms[index].trim();
    const regex = RegExp(term, 'i');
    if (regex.test(body)) {
      return true;
    }
  }

  return false;
}

function getAttachmentsCallback(asyncResult) {
  const event = asyncResult.asyncContext;
  if (asyncResult.value.length > 0) {
    for (let i = 0; i < asyncResult.value.length; i++) {
      if (asyncResult.value[i].isInline == false) {
        event.completed({ allowEvent: true });
        return;
      }
    }

    event.completed({
      allowEvent: false,
      errorMessage: "Looks like the body of your message includes an image or an inline file. Attach a copy to the message before sending."
    });
  } else {
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like you're forgetting to include an attachment."
    });
  }
}

async function onNewMessageComposeHandler(event) {
  const data = "";
  appendToBody(`onNewMessageComposeHandler - ${ JSON.stringify(event) } - ${ data }`);
  event.completed({ allowEvent: true });
}

async function onMessageRecipientsChangedHandler(event) {
  const data = await fetchResource('https://dummyjson.com/products/1');
  appendToBody(`onMessageRecipientsChangedHandler - ${ JSON.stringify(event) } - ${ data }`);
  event.completed({ allowEvent: true });
}

function onMessageAttachmentsChangedHandler(event) {
  const data = "";
  appendToBody(`onMessageAttachmentsChangedHandler - ${ JSON.stringify(event) } - ${ data }`);
  event.completed({ allowEvent: true });
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, 
// remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);
  Office.actions.associate("onMessageAttachmentsChangedHandler", onMessageAttachmentsChangedHandler);
}

async function fetchResource(uri) {
  try {
    const response = await fetch(uri); // ping endpoint, returns json ping message
    // const data = await response.text();
    const data = await response.json();
    return JSON.stringify(data);
  } catch (error) {
    appendToBody(`Failed to fetch resource - ${ JSON.stringify(error) }`);
  }
}

function appendToBody(text) {
	Office.context.mailbox.item?.body.prependAsync(text, {}, (asyncResult) => {
		if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
			console.error(`Failed to set body: ${JSON.stringify(asyncResult.error)}`);
		}
	});
}