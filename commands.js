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

    // 1) Fetch XML (or capture the error)
    let xmlString: string;
    let xmlError: Error | null = null;
    try {
      xmlString = await getCustomXmlPart();
    } catch (err) {
      xmlString = "Error getting custom XML part:";
      xmlError = err as Error;
    }

    // 2) Insert the XML (or error) at the end
    const xmlPara = body.insertParagraph(
      xmlError ? `${xmlString} ${JSON.stringify(xmlError)}` : xmlString,
      Word.InsertLocation.end
    );
    xmlPara.font.color = xmlError ? "#a00" : "#07641d";

    // 3) Force a blank line
    body.insertParagraph("", Word.InsertLocation.end);

    // 4) Insert content control + HTML
    const cc = body.insertContentControl();
    cc.tag = "helloWorldBodyCC";
    cc.title = "Hello World Body Control";
    cc.appearance = Word.ContentControlAppearance.boundingBox;
    cc.insertHtml("<h3>Hello World</h3>", Word.InsertLocation.end);

    // 5) One final sync
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
