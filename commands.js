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

    // 1) Insert your custom XML part (or error) as a paragraph
    let xmlText: string;
    try {
      xmlText = await getCustomXmlPart();
    } catch (err) {
      xmlText = "Error getting custom XML part: " + JSON.stringify(err);
    }
    const xmlPara = body.insertParagraph(xmlText, Word.InsertLocation.end);
    xmlPara.font.color = "#07641d";

    // 2) Blank line for separation
    body.insertParagraph("", Word.InsertLocation.end);

    // 3) Create an *empty* paragraph to host the content control
    const placeholder = body.insertParagraph("", Word.InsertLocation.end);

    // 4) Wrap that paragraph in a content control
    const cc = placeholder.insertContentControl();
    cc.tag = "testHtmlCc";
    cc.title = "Test HTML Content Control";
    cc.appearance = Word.ContentControlAppearance.boundingBox;

    // 5) Now insert your HTML *into* that control, replacing its empty paragraph
    cc.insertHtml(`<h2>Hello World</h2>`, Word.InsertLocation.replace);

    await context.sync();
  });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}

async function paragraphChanged() {
  await Word.run(async (context) => {
    const results = context.document.body.search("110");
    results.load("length");
    await context.sync();
    if (results.items.length == 0) {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.clear();
      header.insertParagraph("Public - The data is for the public and shareable externally", "Start");
      const font = header.font;
      font.color = "#07641d";

      await context.sync();
    } else {
      const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
      header.clear();
      header.insertParagraph("High Confidential - The data must be secret or in some way highly critical", "Start");
      const font = header.font;
      font.color = "#f8334d";
      await context.sync();
    }
  });
}
async function registerOnParagraphChanged(event) {
  Word.run(async (context) => {
    let eventContext = context.document.onParagraphChanged.add(paragraphChanged);
    await context.sync();
  });
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
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

async function getCustomXmlPart() {
  return new Promise<string | undefined>((resolve, reject) => {
    Office.context.document.customXmlParts.getByNamespaceAsync(
      "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
      {},
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
          return;
        }

        const customXmlPart = result.value[0];

        if (!customXmlPart) {
          resolve(undefined);
          return;
        }

        customXmlPart.getXmlAsync({}, (xmlResult) => {
          if (xmlResult.status === Office.AsyncResultStatus.Failed) {
            reject(xmlResult.error);
            return;
          }

          resolve(xmlResult.value);
        });
      }
    );
  });
}

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);
Office.actions.associate("registerOnParagraphChanged", registerOnParagraphChanged);
