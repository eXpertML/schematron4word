/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("RunSchematron").onclick = () => tryCatch(RunSchematron);
  }
});

async function RunSchematron() {
    await Word.run(async (context) => {
        console.log("Getting Word OoXML...");
        const xml = context.document.body.getOoxml();
        await context.sync();
        //console.log(xml.value.toString());
        console.log("Running Schematron...")
        const SVRL = SaxonJS.transform({
          stylesheetLocation: "../../assets/demo.sef.json",
          sourceText: xml.value.toString(),
          destination: "serialized",
            outputProperties: {
              method: "xml",
              indent: false
            }
        });
        await context.sync();
        //console.log(SVRL.principalResult);
        console.log("Transforming to add comments");
        const transformed = SaxonJS.transform({
            stylesheetLocation: "../../assets/InsertSVRLintoWord.sef.json",
            sourceText: xml.value.toString(),
            stylesheetParams: { "SVRLtext": SVRL.principalResult.toString() },
            destination: "serialized"
        });
        await context.sync();
        console.log(transformed.principalResult);    
        console.log("Inserting transformed XML..."); 
        context.document.body.insertOoxml(transformed.principalResult, Word.InsertLocation.replace);
        await context.sync();
        console.log("Done!")
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}
