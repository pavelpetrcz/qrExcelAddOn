import { submitForm } from '../src/taskpane/taskpane.js';

beforeEach(() => {
  // Ensure a clean DOM before each test
  document.body.innerHTML = `
    <button id="generateQrButton">Generate</button>
    <div id="response"></div>
    <input id="bic_prefix" />
    <input id="bic" />
    <select id="bank_code"><option value="0800">0800</option></select>
    <input id="amount" />
    <input id="currency" />
    <input id="vs" />
    <input id="ss" />
    <input id="ks" />
    <input id="msg_input" />
    <input id="qr_dest" />
    <input id="fitToCell" type="checkbox" />
  `;
});

test('submitForm validates required fields and inserts image when Excel available', async () => {
  const btn = document.getElementById('generateQrButton');
  const resp = document.getElementById('response');
  // Set required values
  document.getElementById('bic').value = '0000000000';
  document.getElementById('bank_code').value = '0800';
  document.getElementById('amount').value = '100';
  document.getElementById('qr_dest').value = 'A1';

  // Mock fetch to return a blob
  global.fetch = jest.fn(async () => ({
    ok: true,
    blob: async () => new Blob([new Uint8Array([1,2,3])], { type: 'image/png' })
  }));

  await submitForm();

  expect(global.fetch).toHaveBeenCalled();
  // Excel.run should have been called via our setup mock
  expect(global.Excel.run).toHaveBeenCalled();

  // Response text should indicate insertion success (English string produced in Excel path)
  expect(resp.innerText).toMatch(/inserted successfully|generován úspěšně|úspěšně/i);
  // Button should be enabled at the end
  expect(btn.disabled).toBe(false);
});

test('submitForm shows validation errors when missing required fields', async () => {
  // leave required values empty
  document.getElementById('bic').value = '';
  document.getElementById('bank_code').value = '';
  document.getElementById('amount').value = '';
  document.getElementById('qr_dest').value = '';

  await submitForm();

  const resp = document.getElementById('response');
  expect(resp.innerText).toMatch(/Pole chybí|Číslo účtu|Kód banky|Částka|Cíl QR/);
});
