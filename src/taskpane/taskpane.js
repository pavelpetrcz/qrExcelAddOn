/* global console, document, Excel, Office, window, URLSearchParams, fetch, Uint8Array, btoa */

// --- Global Constant Fix ---
// This constant must be defined for the FRE logic to work.
const FRE_FLAG_KEY = "AddInFreCompleted";
// ---------------------------

// --- Office Initialization and Main Logic ---
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        // 1. Check/Show First-Run Experience
        checkFirstRunStatus();
        
        // 2. Hook up the QR generation button when the add-in is ready
        const btn = document.getElementById("generateQrButton");
        if (btn) {
            btn.addEventListener("click", submitForm);
        }
        console.debug("Office.onReady: Excel host detected. Event bound.");
    }

    // 3. Apply initial theme based on user preference and update when changed
    try {
        const mq = window.matchMedia("(prefers-color-scheme: dark)");
        
        const applyTheme = (e) => {
            // Check e.matches first, fall back to initial mq.matches
            const isDark = (e && typeof e.matches === "boolean") ? e.matches : mq.matches || false;
            document.body.setAttribute("data-theme", isDark ? "dark" : "light");
            console.debug(`Theme applied: ${isDark ? "dark" : "light"}`);
        };
        
        // Use the modern standard addEventListener if available
        if (mq.addEventListener) {
            mq.addEventListener("change", applyTheme);
        // Fallback for older browsers (e.g., specific IE/Edge versions or older Office Webview)
        } else if (mq.addListener) {
            mq.addListener(applyTheme);
        }
        // Apply theme immediately on load
        applyTheme();
    } catch (err) {
        console.debug("Theme detection not available", err);
    }
});

// --- Global Function Exports (Kept for compatibility with HTML module import) ---
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
    
    // Use optional chaining with consistent type assertion (value || "") for safety
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
    const fitToCellCheckbox = document.getElementById("fitToCell"); // Get the element once
    
    const respElm = document.getElementById("response");
    const btn = document.getElementById("generateQrButton");
    
    if (btn) btn.disabled = true;
    if (respElm) respElm.innerText = "Generating…";

    // Basic validation of required inputs
    const missing = [];
    if (!accountNumber || accountNumber.trim() === "") missing.push("Číslo účtu");
    if (!bankCode || bankCode.trim() === "") missing.push("Kód banky");
    if (!amount || amount.trim() === "") missing.push("Částka");
    if (!qr_dest || qr_dest.trim() === "") missing.push("Cílová buňka");

    if (missing.length > 0) {
        if (respElm) respElm.innerText = "Vyplňte prosím pole: " + missing.join(", ");
        if (btn) btn.disabled = false;
        return;
    }
    
    // Amount validation and conversion
    // Using parseFloat instead of Number() for better handling of decimal inputs
    const amountNum = parseFloat(amount.toString().replace(",", ".")); 
    
    if (Number.isNaN(amountNum) || amountNum <= 0) {
        if (respElm) respElm.innerText = "Neplatná částka. Zadejte prosím kladné číslo."; // Changed text to reflect floating point support
        if (btn) btn.disabled = false;
        return;
    }

    // Define the API endpoint URL (HTTPS is correct)
    const apiUrl = "https://api.paylibo.com/paylibo/generator/czech/image";

    // Build the query string with only non-empty parameters
    const queryParams = new URLSearchParams();
    if (accountPrefix) queryParams.set("accountPrefix", accountPrefix);
    if (accountNumber) queryParams.set("accountNumber", accountNumber);
    if (bankCode) queryParams.set("bankCode", bankCode);
    // Use toFixed(2) to ensure two decimal places for the API, as amountNum is a float
    queryParams.set("amount", amountNum.toFixed(2)); 
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
            // Include status for better debugging
            throw new Error(`Server returned status ${response.status}. Vstupní data QR kódu nejsou v pořádku. Zkontrolujte je prosím. A zkuste to znovu.`);
        }
        const blob = await response.blob();

        // Convert blob/arrayBuffer to base64
        const arrayBuffer = await blob.arrayBuffer();
        const bytes = new Uint8Array(arrayBuffer);
        
        // Optimized ArrayBuffer to Base64 conversion
        // NOTE: The previous `String.fromCharCode.apply(null, chunk)` method can fail 
        // on large arrays due to argument list limits. Using map/join is safer.
        let binary = "";
        for (let i = 0; i < bytes.byteLength; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        const base64String = btoa(binary);

        // Insert image into Excel in the specified range (if available)
        if (typeof Excel !== "undefined" && Excel.run) {
            await Excel.run(async (context) => {
                try {
                    const worksheet = context.workbook.worksheets.getActiveWorksheet();
                    const range = worksheet.getRange(qr_dest);
                    range.load(["left", "top", "width", "height"]);
                    await context.sync();

                    // Add the image using the base64 string
                    const image = worksheet.shapes.addImage(base64String);
                    image.left = range.left;
                    image.top = range.top;
                    
                    // Check the checkbox state
                    const fitToCell = fitToCellCheckbox && fitToCellCheckbox.checked;
                    
                    if (fitToCell) {
                        image.width = range.width;
                        image.height = range.height;
                    }
                    
                    await context.sync();
                    console.log("Image inserted successfully");
                    if (respElm) respElm.innerText = "QR kód úspěšně vložen!";
                } catch (err) {
                    // Provide a cleaner error message
                    const errorMessage = err.message || "Unknown error";
                    if (respElm) respElm.innerText = `Chyba vložení QR kódu. Zkontrolujte prosím buňku (${qr_dest}). Detaily: ${errorMessage}`;
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
                // Clear the 'Generating...' text and append the image
                respElm.innerHTML = ""; 
                respElm.appendChild(previewImg);
                
            }
            if (respElm) respElm.innerText += " QR kód vygenerován úspěšně (náhled v prohlížeči).";
        }
    } catch (error) {
        if (respElm) respElm.innerText = "Chyba načtení QR kódu: " + (error && error.message);
        console.error("Error fetching image:", error);
    } finally {
        if (btn) btn.disabled = false;
    }
}

// --- First-Run Experience Logic ---

// NOTE: This listener is redundant if you are using the Office.onReady listener,
// but it is kept to support development outside of the Office host environment.
document.addEventListener("DOMContentLoaded", () => {
    // Check if Office.js has *not* loaded yet (i.e., not in an Office host)
    // The button listener is already added in Office.onReady for the production path.
    if (typeof Office === "undefined" || !Office.context) { 
        const btn = document.getElementById("generateQrButton");
        if (btn) btn.addEventListener("click", submitForm);
        console.debug("DOMContentLoaded: Event bound for button (Non-Office host).");
    }
});

function checkFirstRunStatus() {
    const freContainer = document.getElementById("fre-container");
    const mainContainer = document.getElementById("main-container");
    const startButton = document.getElementById("start-button");
    
    // Check for null elements for robustness
    if (!freContainer || !mainContainer || !startButton) {
        console.warn("FRE elements not found. Skipping first-run check.");
        return;
    }

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
    
    if (freContainer && mainContainer) {
        freContainer.classList.add("hidden");
        mainContainer.classList.remove("hidden");
        console.log("FRE completed. Transitioning to main add-in content.");
    }
}