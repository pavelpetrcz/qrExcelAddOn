
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Hook up the QR generation button when the add-in is ready
    const btn = document.getElementById('generateQrButton');
    if (btn) {
      btn.addEventListener('click', submitForm);
    }
    console.debug('Office.onReady: Excel host detected. Event bound.');
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      // Insert your Excel code here
      const range = context.workbook.getSelectedRange();
      range.load('address');

      // Update the fill color
      range.format.fill.color = 'yellow';

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

export async function submitForm() {
  console.debug('submitForm invoked');
  const accountPrefix = document.getElementById('bic_prefix')?.value || '';
  const accountNumber = document.getElementById('bic')?.value || '';
  const bankCode = document.getElementById('bank_code')?.value || '';
  const amount = document.getElementById('amount')?.value || '';
  const currency = document.getElementById('currency')?.value || '';
  const vs = document.getElementById('vs')?.value || '';
  const ss = document.getElementById('ss')?.value || '';
  const ks = document.getElementById('ks')?.value || '';
  const message = document.getElementById('msg_input')?.value || '';
  const qr_dest = document.getElementById('qr_dest')?.value || '';

  const respElm = document.getElementById('response');
  const btn = document.getElementById('generateQrButton');
  if (btn) btn.disabled = true;
  if (respElm) respElm.innerText = 'Generating…';

  // Basic validation
  if (!qr_dest) {
    if (respElm) respElm.innerText = 'Please provide a target cell (e.g., E2) into which the QR image will be inserted.';
    if (btn) btn.disabled = false;
    return;
  }

  // Define the API endpoint URL (use HTTPS to avoid mixed-content errors)
  const apiUrl = 'https://api.paylibo.com/paylibo/generator/czech/image';

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
    message,
  });

  const fullUrl = `${apiUrl}?${queryString.toString()}`;

  try {
    const response = await fetch(fullUrl);
    if (!response.ok) {
      throw new Error('Network response was not ok');
    }
    const blob = await response.blob();

    // Convert blob/arrayBuffer to base64
    const arrayBuffer = await blob.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }
    const base64String = btoa(binary);

    // Insert image into Excel in the specified range (if available), otherwise show preview
    if (typeof Excel !== 'undefined' && Excel.run) {
      await Excel.run(async (context) => {
        try {
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          const range = worksheet.getRange(qr_dest);
          range.load(['left', 'top', 'width', 'height']);
          await context.sync();

          // Add the image using the base64 string (no data URL prefix)
          const image = worksheet.shapes.addImage(base64String);
          image.left = range.left;
          image.top = range.top;
          image.width = range.width;
          image.height = range.height;
          await context.sync();
          console.log('Image inserted successfully');
          if (respElm) respElm.innerText = 'QR image generated and inserted successfully.';
        } catch (err) {
          if (respElm) respElm.innerText = 'Error inserting image: ' + (err && err.message);
          console.error('Error inserting image:', err);
        }
      });
    } else {
      // Browser preview: show data URL image for development/testing
      const previewImg = document.createElement('img');
      previewImg.src = `data:image/png;base64,${base64String}`;
      previewImg.style.maxWidth = '150px';
      previewImg.style.maxHeight = '150px';
      if (respElm) {
        respElm.innerHTML = '';
        respElm.appendChild(previewImg);
      }
      if (respElm) respElm.innerText = 'QR generated (preview).';
    }
  } catch (error) {
    if (respElm) respElm.innerText = 'Error fetching image: ' + (error && error.message);
    console.error('Error fetching image:', error);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Attach listener for plain DOM / browser testing (not running inside Excel)
document.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('generateQrButton');
  if (btn) btn.addEventListener('click', submitForm);
  console.debug('DOMContentLoaded: Event bound for button');
});
