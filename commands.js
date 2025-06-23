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

      let xml = '';
      try {
        xml = await getCustomXmlPart();
        header.insertParagraph(xml, "Start");
      } catch(error) {
         header.insertParagraph("some error" + JSON.stringify(error), "Start");
      }
    
      header.font.color = "#07641d";
      await context.sync();
  });

  await insertContentControlHtml();

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
    }
    else {
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

async function insertContentControlHtml() {
  await Word.run(async (context) => {
     const firstSection = context.document.sections.getFirst();
    // 2. Get its Primary header
    const header = firstSection.getHeader("Primary");
    // 3. Get the full range of that header
    const headerRange = header.getRange();
    
    // 4. Wrap that range in a content control
    const cc = headerRange.insertContentControl();
    cc.tag = "helloWorldHeaderCC";
    cc.title = "Hello World Header Control";
    cc.appearance = "BoundingBox";
    
    // 5. Insert your HTML inside the content control
    cc.insertHtml("<h3>Hello World</h3>", Word.InsertLocation.replace);
    
    await context.sync();
  })
  .catch((error) => {
    console.error("Error inserting HTML:", error);
  });
}

// The add-in command functions need to be available in global scope

Office.actions.associate("changeHeader", changeHeader);
Office.actions.associate("registerOnParagraphChanged", registerOnParagraphChanged);
