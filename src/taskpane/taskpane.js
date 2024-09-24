/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
      document.getElementById("selectTextButton").onclick = selectText;
  }
});

function selectText() {
  Word.run((context) => {
      const selection = context.document.getSelection();
      selection.load('text');
      return context.sync().then(() => {
          const selectedText = selection.text;
          openDialog(selectedText);
      });
  })
  .catch((error) => {
      console.error(error);
  });
}

function openDialog(selectedText) {
  const url = 'dialog.html'; // Path to your dialog HTML
  Office.context.ui.displayDialogAsync(url, { height: 50, width: 50 }, (asyncResult) => {
      const dialog = asyncResult.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          console.log(arg.message);
      });
      dialog.messageParent(selectedText); // Send selected text to dialog
  });
}

