/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// The initialize function must be run each time a new page is loaded
Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("runFunctions").onclick = runFunctions;
  document.getElementById("copyWorkbook").onclick = copyWorkbook;
});

export async function runFunctions() {
  try {
    await Excel.run(async (context) => {
      let activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
      activeWorksheet.calculate(true);
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function copyWorkbook() {
  try {
    await Excel.run(async (context) => {
      getDocumentAsCompressed();
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// The following example gets the document in Office Open XML ("compressed") format in 65536 bytes (64 KB) slices.
function getDocumentAsCompressed() {
  document.getElementById("status").value = "start";

  Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 /*64 KB*/ }, function (result) {
    if (result.status == "succeeded") {
      // If the getFileAsync call succeeded, then
      // result.value will return a valid File Object.
      const myFile = result.value;
      const sliceCount = myFile.sliceCount;
      const docdataSlices = [];
      let slicesReceived = 0,
        gotAllSlices = true;

      document.getElementById("status").value = "File size:" + myFile.size + " #Slices: " + sliceCount;

      // Get the file slices.
      getSliceAsync(myFile, 0, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
    } else {
      document.getElementById("status").value = "Error:" + result.error.message;
    }
  });
}

function getSliceAsync(file, nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived) {
  file.getSliceAsync(nextSlice, function (sliceResult) {
    if (sliceResult.status == "succeeded") {
      if (!gotAllSlices) {
        // Failed to get all slices, no need to continue.
        return;
      }

      // Got one slice, store it in a temporary array.
      docdataSlices[sliceResult.value.index] = sliceResult.value.data;
      if (++slicesReceived == sliceCount) {
        // All slices have been received.
        file.closeAsync();
        onGotAllSlices(docdataSlices);
      } else {
        getSliceAsync(file, ++nextSlice, sliceCount, gotAllSlices, docdataSlices, slicesReceived);
      }
    } else {
      gotAllSlices = false;
      file.closeAsync();
      document.getElementById("status").value = "getSliceAsync Error:" + sliceResult.error.message;
    }
  });
}

function onGotAllSlices(docdataSlices) {
  let docdata = [];
  for (let i = 0; i < docdataSlices.length; i++) {
    docdata = docdata.concat(docdataSlices[i]);
  }

  let fileContent = new String();
  for (let j = 0; j < docdata.length; j++) {
    fileContent += String.fromCharCode(docdata[j]);
  }

  // Now all the file content is stored in 'fileContent' variable
  document.getElementById("result").value = fileContent;

  async function newWorkbookFromFile() {
    await Excel.createWorkbook(fileContent);
  }

  newWorkbookFromFile();
}
