/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("generateQrButton").onclick = submitForm;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}




function submitForm() {
  const accountPrefix = document.getElementById('bic_prefix').value;
  const accountNumber = document.getElementById('bic').value;
  const bankCode = document.getElementById('bank_code').value;
  const amount = document.getElementById('amount').value;
  const currency = document.getElementById('currency').value;
  const vs = document.getElementById('vs').value;
  const ss = document.getElementById('ss').value;
  const ks = document.getElementById('ks').value;
  const message = document.getElementById('msg_input').value;
  const qr_dest = document.getElementById('qr_dest').value;

  // Define the API endpoint URL
  const apiUrl = 'http://api.paylibo.com/paylibo/generator/czech/image';

  // Build the query string with parameters
  const queryString = new URLSearchParams({
      accountPrefix,
      accountNumber,
      bankCode,
      amount,
      currency,
      vs,
      ks,
      ss,
      message
  });

  // Construct the full URL with query string
  const fullUrl = `${apiUrl}?${queryString.toString()}`;

  // Make the API call using Fetch API
  try {
      fetch(fullUrl)
      if (!response.ok) {
          throw new Error('Network response was not ok');
      }
      const blob = response.blob();
      const
          objectUrl = URL.createObjectURL(blob);

      const img = document.createElement('img');
      img.src = objectUrl;
      document.body.appendChild(img); // Replace with your desired container
  } catch (error) {
      console.error('Error fetching image:', error);
  }

  Excel.run(async (context) => {
      try {
          const imageData = new Uint8Array(img);

          const worksheet = context.workbook.worksheets.getItem(sheetName);
          const shapes = worksheet.shapes;
          const range = worksheet.getRange(qr_dest);

          const picture = await shapes.addPictureAsync({
              data: imageData,
              left: range.left,
              top: range.top,
              width: range.width,
              height: range.height
          });

          await context.sync();
          console.log('Image inserted successfully');
      } catch (error) {
          console.error('Error inserting image:', error);
      }
  });

}
