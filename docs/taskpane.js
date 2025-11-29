/* Copy of src/taskpane/taskpane.js for GitHub Pages */
/* Minimalized for deploy */
/* global console, document, Excel, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const btn = document.getElementById('generateQrButton');
    if (btn) btn.addEventListener('click', submitForm);
  }
});

export async function run() { /* intentionally minimal */ }

export async function submitForm() {
  const accountPrefix = document.getElementById('bic_prefix')?.value || '';
  const accountNumber = document.getElementById('bic')?.value || '';
  const bankCode = document.getElementById('bank_code')?.value || '';
  const amount = document.getElementById('amount')?.value || '';
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
  const errors = [];
  if (!accountNumber) errors.push('Číslo účtu');
  if (!bankCode) errors.push('Kód banky');
  if (!amount) errors.push('Částka');
  if (!qr_dest) errors.push('Cíl QR');
  if (errors.length) { if (respElm) respElm.innerText = 'Missing: ' + errors.join(', '); if (btn) btn.disabled = false; return; }

  const apiUrl = 'https://api.paylibo.com/paylibo/generator/czech/image';
  const params = new URLSearchParams();
  if (accountPrefix) params.set('accountPrefix', accountPrefix);
  if (accountNumber) params.set('accountNumber', accountNumber);
  if (bankCode) params.set('bankCode', bankCode);
  params.set('amount', amount);
  if (vs) params.set('vs', vs);
  if (ss) params.set('ss', ss);
  if (ks) params.set('ks', ks);
  if (message) params.set('message', message);
  const fullUrl = `${apiUrl}?${params.toString()}`;

  try {
    const response = await fetch(fullUrl);
    const blob = await response.blob();
    const arrayBuffer = await blob.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = '';
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }
    const base64String = btoa(binary);
    const previewImg = document.createElement('img');
    previewImg.src = `data:image/png;base64,${base64String}`;
    previewImg.style.maxWidth = '150px';
    previewImg.style.maxHeight = '150px';
    if (respElm) { respElm.innerHTML = ''; respElm.appendChild(previewImg); }
  } catch (err) {
    if (respElm) respElm.innerText = 'Error: ' + (err && err.message);
  } finally {
    if (btn) btn.disabled = false;
  }
}

document.addEventListener('DOMContentLoaded', () => { const btn = document.getElementById('generateQrButton'); if (btn) btn.addEventListener('click', submitForm); });
