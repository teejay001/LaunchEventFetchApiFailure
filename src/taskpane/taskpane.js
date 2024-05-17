/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

async function fetchResource(uri) {
  try {
    const response = await fetch(uri); // ping endpoint, returns json ping message
    // const data = await response.text();
    const data = await response.json();
    return JSON.stringify(data);
  } catch (error) {
    appendToBody(`Failed to fetch resource - ${ JSON.stringify(error) }`);
  }
};

function appendToBody(text) {
	Office.context.mailbox.item?.body.prependAsync(text, {}, (asyncResult) => {
		if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
			console.error(`Failed to set body: ${JSON.stringify(asyncResult.error)}`);
		}
	});
};

export async function run() {
  const data = await fetchResource('https://dummyjson.com/products/1');
  appendToBody(`from taskpane Run - ${ data }`);

  // // Get a reference to the current message
  // const item = Office.context.mailbox.item;

  // // Write message property value to the task pane
  // document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
}
