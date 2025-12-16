/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    checkFirstRunStatus();
    // Hook up the QR generation button when the add-in is ready
    const btn = document.getElementById("generateQrButton");
    if (btn) {
      btn.addEventListener("click", submitForm);
    }
    console.debug("Office.onReady: Excel host detected. Event bound.");
  }

  // Apply initial theme based on user preference and update when changed
  try {
    const mq = window.matchMedia("(prefers-color-scheme: dark)");
    const applyTheme = (e) => {
      const isDark = e && typeof e.matches !== "undefined" ? e.matches : mq.matches || false;
      document.body.setAttribute("data-theme", isDark ? "dark" : "light");
    };
    // Add event listener if possible
    if (mq.addEventListener) {
      mq.addEventListener("change", applyTheme);
    } else if (mq.addListener) {
      mq.addListener(applyTheme);
    }
    applyTheme();
  } catch (err) {
    console.debug("Theme detection not available", err);
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      // Insert your Excel code here
      const range = context.workbook.getSelectedRange();
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

export async function submitForm() {
  console.debug("submitForm invoked");
  const accountPrefix = document.getElementById("bic_prefix")?.value || "";
  const accountNumber = document.getElementById("bic")?.value || "";
  const bankCode = document.getElementById("bank_code")?.value || "";
  const amount = document.getElementById("amount")?.value || "";
  const currency = document.getElementById("currency")?.value || "";
  const vs = document.getElementById("vs")?.value || "";
  const ss = document.getElementById("ss")?.value || "";
  const ks = document.getElementById("ks")?.value || "";
  const message = document.getElementById("msg_input")?.value || "";
  const qr_dest = document.getElementById("qr_dest")?.value || "";

  const respElm = document.getElementById("response");
  const btn = document.getElementById("generateQrButton");
  if (btn) btn.disabled = true;
  if (respElm) respElm.innerText = "Generating…";

  // Basic validation of required inputs
  const missing = [];
  // accountPrefix is optional now; do not collect it in missing
  if (!accountNumber || accountNumber.trim() === "") missing.push("Číslo účtu");
  if (!bankCode || bankCode.trim() === "") missing.push("Kód banky");
  if (!amount || amount.trim() === "") missing.push("Částka");
  if (!qr_dest || qr_dest.trim() === "") missing.push("Cílová buňka");

  if (missing.length > 0) {
    if (respElm) respElm.innerText = "Vyplňte prosím pole: " + missing.join(", ");
    if (btn) btn.disabled = false;
    return;
  }
  // amount validation
  const amountNum = Number(amount.toString().replace(",", "."));
  if (Number.isNaN(amountNum) || amountNum <= 0) {
    if (respElm) respElm.innerText = "Neplatná částka. Zadejte prosím kladné celé číslo.";
    if (btn) btn.disabled = false;
    return;
  }

  // Define the API endpoint URL (use HTTPS to avoid mixed-content errors)
  const apiUrl = "https://api.paylibo.com/paylibo/generator/czech/image";

  // Build the query string with only non-empty parameters
  const queryParams = new URLSearchParams();
  if (accountPrefix) queryParams.set("accountPrefix", accountPrefix);
  if (accountNumber) queryParams.set("accountNumber", accountNumber);
  if (bankCode) queryParams.set("bankCode", bankCode);
  if (amountNum) queryParams.set("amount", amountNum.toString());
  if (currency) queryParams.set("currency", currency);
  if (vs) queryParams.set("vs", vs);
  if (ks) queryParams.set("ks", ks);
  if (ss) queryParams.set("ss", ss);
  if (message) queryParams.set("message", message);

  const fullUrl = `${apiUrl}?${queryParams.toString()}`;
  console.debug("Full API URL will be called:", fullUrl);

  try {
    const response = await fetch(fullUrl);
    if (!response.ok) {
      throw new Error("Vstupní data QR kódu nejsou v pořádku. Zkontrolujte je prosím. A zkuste to znovu.");
    }
    const blob = await response.blob();

    // Convert blob/arrayBuffer to base64
    const arrayBuffer = await blob.arrayBuffer();
    const bytes = new Uint8Array(arrayBuffer);
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk);
    }
    const base64String = btoa(binary);

    // Insert image into Excel in the specified range (if available), otherwise show preview
    if (typeof Excel !== "undefined" && Excel.run) {
      await Excel.run(async (context) => {
        try {
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          const range = worksheet.getRange(qr_dest);
          range.load(["left", "top", "width", "height"]);
          await context.sync();

          // Add the image using the base64 string (no data URL prefix)
          // Respect the user's choice to resize or not:
          const image = worksheet.shapes.addImage(base64String);
          image.left = range.left;
          image.top = range.top;
          // If the checkbox is checked, resize to cell, otherwise keep native image size
          const fitToCell = document.getElementById("fitToCell") && document.getElementById("fitToCell").checked;
          if (fitToCell) {
            image.width = range.width;
            image.height = range.height;
          }
          await context.sync();
          console.log("Image inserted successfully");
          if (respElm) respElm.innerText = "QR kód úspěšně vložen!";
        } catch (err) {
          if (respElm) respElm.innerText = "Chyba vložení QR kódu: " + (err && err.message);
          console.error("Error inserting image:", err);
        }
      });
    } else {
      // Browser preview: show data URL image for development/testing
      const previewImg = document.createElement("img");
      previewImg.src = `data:image/png;base64,${base64String}`;
      previewImg.style.maxWidth = "150px";
      previewImg.style.maxHeight = "150px";
      if (respElm) {
        respElm.innerHTML = "";
        respElm.appendChild(previewImg);
      }
      if (respElm) respElm.innerText = "QR kód vygenerován úspěšně.";
    }
  } catch (error) {
    if (respElm) respElm.innerText = "Chyba načtení QR kódu: " + (error && error.message);
    console.error("Error fetching image:", error);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// Attach listener for plain DOM / browser testing (not running inside Excel)
document.addEventListener("DOMContentLoaded", () => {
  const btn = document.getElementById("generateQrButton");
  if (btn) btn.addEventListener("click", submitForm);
  console.debug("DOMContentLoaded: Event bound for button");
});

function checkFirstRunStatus() {
  const freContainer = document.getElementById("fre-container");
  const mainContainer = document.getElementById("main-container");
  const startButton = document.getElementById("start-button");

  // 1. Check if the flag exists in localStorage
  const hasCompletedFRE = localStorage.getItem(FRE_FLAG_KEY);

  if (hasCompletedFRE === "true") {
    // The flag is set: Show the main add-in UI
    freContainer.classList.add("hidden");
    mainContainer.classList.remove("hidden");
  } else {
    // The flag is NOT set: Show the First-Run Experience UI
    freContainer.classList.remove("hidden");
    mainContainer.classList.add("hidden");

    // Attach event listener to the "Start" button
    startButton.addEventListener("click", completeFirstRun);
  }
}

function completeFirstRun() {
  // 2. Set the flag in localStorage
  localStorage.setItem(FRE_FLAG_KEY, "true");

  // 3. Transition the view (hide FRE, show main content)
  const freContainer = document.getElementById("fre-container");
  const mainContainer = document.getElementById("main-container");

  freContainer.classList.add("hidden");
  mainContainer.classList.remove("hidden");

  // Optional: Log the transition
  console.log("FRE completed. Transitioning to main add-in content.");
}
