/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

async function changeHeader(event) {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
    const firstPageHeader = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.firstPage);
    header.clear();
    firstPageHeader.clear();
    const xmlString = await getCustomXmlPart();
    header.insertParagraph(`${xmlString}`, "Start");
    header.font.color = "#07641d";

    await context.sync();
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

function getCustomXmlPart() {
  return Office.context.document.customXmlParts.getByNamespaceAsync(
    "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    {},
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.log("Error retrieving custom XML part: " + result.error.message);
        return;
      }

      const customXmlPart = result.value[0];

      if (!customXmlPart) {
        console.log("No custom XML part found with the specified namespace.");
        return;
      }

      return customXmlPart.getXmlAsync({}, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          return;
        }

        return result.value;
      });
    }
  );
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);
