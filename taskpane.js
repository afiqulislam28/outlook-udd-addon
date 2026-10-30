/**
 * UDD Downgrade Sync — Outlook Add-in
 * Links: Outlook Desktop (Windows / Mac), Outlook on the Web (Office 365), and Outlook Mobile (Android/iOS)
 * Target Destination: Downgrade sheet in "DISTRIBUTOR REQUEST - AGM PLANNING.xlsx" via Power Automate
 */

// Default Power Automate Bridge Endpoint
const DEFAULT_BRIDGE_URL = "https://defaultbda5f9dab8cd47efba3f91645269a2.48.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/2ab50687df5a4a978c31bf38d705121c/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=v45QYDIZX08hHSRwb31odqYWiK1RyEeVrl3APaf7fnw";

// NTTN Rate Slabs for SCL and FHL
const SLABS = {
  SCL: {
    D2D: [
      { min: 0, max: 11, rate: 70 },
      { min: 12, max: 99, rate: 60 },
      { min: 100, max: 299, rate: 25 },
      { min: 300, max: 100000, rate: 12 }
    ],
    LH: [
      { min: 0, max: 11, rate: 80 },
      { min: 12, max: 99, rate: 70 },
      { min: 100, max: 299, rate: 25 },
      { min: 300, max: 499, rate: 21 },
      { min: 500, max: 999, rate: 20 },
      { min: 1000, max: 2499, rate: 18 },
      { min: 2500, max: 9999, rate: 16 },
      { min: 10000, max: 100000, rate: 15 }
    ],
    METRO: [
      { min: 0, max: 11, rate: 60 },
      { min: 12, max: 99, rate: 55 },
      { min: 100, max: 299, rate: 15 },
      { min: 300, max: 100000, rate: 11 }
    ]
  },
  FHL: {
    D2D: [
      { min: 11, max: 50, rate: 90 },
      { min: 51, max: 100, rate: 35 },
      { min: 101, max: 299, rate: 18 },
      { min: 300, max: 999, rate: 14 },
      { min: 1000, max: 100000, rate: 12.5 }
    ],
    LH: [
      { min: 11, max: 50, rate: 115 },
      { min: 51, max: 100, rate: 55 },
      { min: 101, max: 299, rate: 25 },
      { min: 300, max: 499, rate: 24 },
      { min: 500, max: 100000, rate: 23 }
    ],
    METRO: [
      { min: 11, max: 50, rate: 70 },
      { min: 51, max: 100, rate: 35 },
      { min: 101, max: 299, rate: 18 },
      { min: 300, max: 2499, rate: 13 },
      { min: 2500, max: 100000, rate: 12.5 }
    ]
  }
};

/**
 * Calculates NTTN rate based on vendor slabs + 5% VAT
 */
function calculateVendorRate(vendor, category, bw) {
  if (!vendor || !category || !bw || !SLABS[vendor] || !SLABS[vendor][category]) return "0.00";
  const vSlabs = SLABS[vendor][category];
  const bwNum = parseFloat(bw) || 0;
  let match = vSlabs.find(r => bwNum >= r.min && bwNum <= r.max);
  let rate = match ? match.rate : 0;

  if (vendor === "FHL" && bwNum > 0) {
    if (category === "D2D") rate = Math.max(rate, 1650 / bwNum);
    else if (category === "LH") rate = Math.max(rate, 2200 / bwNum);
    else if (category === "METRO") rate = Math.max(rate, 1500 / bwNum);
  }
  return (rate * 1.05).toFixed(2);
}

/**
 * Normalizes vendor name to standard short form
 */
function shortenVendor(name) {
  if (!name) return "SCL";
  const upper = name.toUpperCase();
  if (upper.includes("SUMMIT") || upper.includes("SCL")) return "SCL";
  if (upper.includes("FIBER") || upper.includes("FHL")) return "FHL";
  if (upper.includes("BAHON")) return "Bahon";
  if (upper.includes("BTCL") || upper.includes("BANGLADESH TELECOM")) return "BTCL";
  if (upper.includes("EARTH") || upper.includes("EART")) return "Earth";
  return name.trim();
}

// Global State
let currentEmail = {
  subject: "",
  from: "",
  date: "",
  rawDate: new Date(),
  bodyHtml: "",
  bodyText: "",
  records: []
};

// Initialize Office.js
Office.onReady((info) => {
  console.log("Office.js initialized. Host:", info.host, "Platform:", info.platform);

  // Platform badge
  const platformBadge = document.getElementById("platform-badge");
  if (platformBadge) {
    let p = "Outlook";
    if (info.platform === Office.PlatformType.PC) p = "Outlook Desktop (Windows)";
    else if (info.platform === Office.PlatformType.Mac) p = "Outlook Desktop (Mac)";
    else if (info.platform === Office.PlatformType.OfficeOnline) p = "Outlook Web (365)";
    else if (info.platform === Office.PlatformType.Android) p = "Outlook Android";
    else if (info.platform === Office.PlatformType.iOS) p = "Outlook iOS";
    platformBadge.textContent = p;
  }

  // Load Bridge URL from storage or default
  initSettings();

  // Load email content
  loadEmailData();

  // Listen for email selection change when taskpane is pinned
  if (Office.context && Office.context.mailbox && Office.context.mailbox.addHandlerAsync) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      console.log("Email selection changed, refreshing...");
      loadEmailData();
    });
  }

  // UI Event Listeners
  document.getElementById("btn-sync-sheet")?.addEventListener("click", handleSyncClick);
  document.getElementById("btn-copy-tsv")?.addEventListener("click", handleCopyTsvClick);
  document.getElementById("btn-refresh-email")?.addEventListener("click", loadEmailData);
  document.getElementById("btn-add-record")?.addEventListener("click", handleAddManualRecord);
  document.getElementById("btn-save-url")?.addEventListener("click", saveBridgeUrl);
  document.getElementById("btn-reset-url")?.addEventListener("click", resetBridgeUrl);
});

function initSettings() {
  const input = document.getElementById("bridge-url-input");
  const saved = localStorage.getItem("udd_bridge_url");
  if (input) {
    input.value = saved || DEFAULT_BRIDGE_URL;
  }
}

function getActiveBridgeUrl() {
  return localStorage.getItem("udd_bridge_url") || DEFAULT_BRIDGE_URL;
}

function saveBridgeUrl() {
  const input = document.getElementById("bridge-url-input");
  if (input && input.value.trim()) {
    localStorage.setItem("udd_bridge_url", input.value.trim());
    showStatus("Power Automate Bridge URL saved successfully!", "success");
  }
}

function resetBridgeUrl() {
  const input = document.getElementById("bridge-url-input");
  if (input) {
    input.value = DEFAULT_BRIDGE_URL;
    localStorage.removeItem("udd_bridge_url");
    showStatus("Bridge URL reset to default.", "info");
  }
}

/**
 * Loads metadata and body from currently open email
 */
function loadEmailData() {
  const item = Office.context.mailbox.item;
  if (!item) {
    showStatus("Please open an email message in Outlook.", "error");
    return;
  }

  // UI state
  document.getElementById("loading-state").style.display = "flex";
  document.getElementById("main-content").style.display = "none";
  hideStatus();

  currentEmail.subject = item.subject || "No Subject";
  currentEmail.from = item.from ? `${item.from.displayName} (${item.from.emailAddress})` : "Unknown Sender";
  currentEmail.rawDate = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
  currentEmail.date = currentEmail.rawDate.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });

  document.getElementById("email-subject").textContent = currentEmail.subject;
  document.getElementById("email-from").textContent = currentEmail.from;
  document.getElementById("email-date").textContent = currentEmail.date;

  // Fetch email body
  item.body.getAsync(Office.CoercionType.Html, (htmlRes) => {
    if (htmlRes.status === Office.AsyncResultStatus.Succeeded && htmlRes.value) {
      currentEmail.bodyHtml = htmlRes.value;
      parseEmailContent(htmlRes.value, "");
    } else {
      // Fallback to text
      item.body.getAsync(Office.CoercionType.Text, (txtRes) => {
        const text = txtRes.status === Office.AsyncResultStatus.Succeeded ? txtRes.value : "";
        currentEmail.bodyText = text;
        parseEmailContent("", text);
      });
    }

    document.getElementById("loading-state").style.display = "none";
    document.getElementById("main-content").style.display = "block";
  });
}

/**
 * Parses email HTML or text body to extract capacity downgrade tables
 */
function parseEmailContent(htmlContent, textContent) {
  let records = [];

  if (htmlContent) {
    records = parseHtmlTables(htmlContent);
  }

  if (records.length === 0 && (textContent || htmlContent)) {
    const rawText = textContent || stripHtml(htmlContent);
    records = parseTextBody(rawText);
  }

  // If still empty, create an initial blank record pre-filled with context from subject
  if (records.length === 0) {
    const guess = guessFromSubject(currentEmail.subject);
    records.push({
      included: true,
      linkId: guess.linkId || "",
      distri: guess.distri || "",
      vendor: guess.vendor || "SCL",
      existingBw: "",
      downgradeBw: "",
      newBw: "",
      category: "D2D",
      unitCost: "0.00",
      effectiveDate: currentEmail.date || "ASAP",
      mathNote: "Manual input (no table auto-detected)",
      isMathValid: true
    });
    updateStatusBadge("Manual Entry", "warning");
  } else {
    updateStatusBadge(`${records.length} Link(s) Detected`, "success");
  }

  currentEmail.records = records;
  renderRecordsUI();
}

/**
 * Searches HTML tables for Link ID, Existing BW, Downgrade BW, New BW
 */
function parseHtmlTables(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const tables = doc.querySelectorAll("table");
  const results = [];

  for (let table of tables) {
    const rows = table.querySelectorAll("tr");
    if (rows.length < 2) continue;

    // Header inspection
    let headerIdx = -1;
    let colMap = {};

    for (let rIdx = 0; rIdx < rows.length; rIdx++) {
      const cells = rows[rIdx].querySelectorAll("th, td");
      const rowTexts = Array.from(cells).map(c => c.textContent.trim().toLowerCase());
      
      const hasLink = rowTexts.some(c => c.includes("link"));
      const hasCap = rowTexts.some(c => c.includes("cap") || c.includes("bw") || c.includes("bandwidth"));

      if (hasLink || hasCap) {
        headerIdx = rIdx;
        rowTexts.forEach((col, cIdx) => {
          if (col.includes("link")) colMap["linkId"] = cIdx;
          else if (col.includes("current") || col.includes("existing") || col.includes("old") || col.includes("present")) colMap["existing"] = cIdx;
          else if ((col.includes("down") || col.includes("decrease")) && !col.includes("new") && !col.includes("after")) colMap["downgrade"] = cIdx;
          else if (col.includes("new") || col.includes("after") || col.includes("revised")) colMap["new"] = cIdx;
          else if (col.includes("provider") || col.includes("vendor") || col.includes("nttn")) colMap["vendor"] = cIdx;
          else if (col.includes("bts") || col.includes("distributor") || col.includes("service agent") || col.includes("client") || col.includes("sa name")) colMap["distri"] = cIdx;
          else if (col.includes("effective") || col.includes("date")) colMap["effectiveDate"] = cIdx;
        });
        break;
      }
    }

    if (headerIdx !== -1 && (colMap["linkId"] !== undefined || colMap["downgrade"] !== undefined)) {
      for (let rIdx = headerIdx + 1; rIdx < rows.length; rIdx++) {
        const cells = rows[rIdx].querySelectorAll("td, th");
        if (cells.length === 0) continue;

        const getVal = (idx) => idx !== undefined && cells[idx] ? cells[idx].textContent.trim() : "";

        let linkId = getVal(colMap["linkId"]);
        let existingStr = getVal(colMap["existing"]);
        let downStr = getVal(colMap["downgrade"]);
        let newStr = getVal(colMap["new"]);
        let vendorStr = getVal(colMap["vendor"]);
        let distriStr = getVal(colMap["distri"]);
        let dateStr = getVal(colMap["effectiveDate"]);

        // Ignore summary/total rows
        const rowText = rows[rIdx].textContent.toLowerCase();
        if (rowText.includes("total") || rowText.includes("regard") || rowText.includes("thank")) continue;

        let existingNum = extractNumeric(existingStr);
        let downNum = extractNumeric(downStr);
        let newNum = extractNumeric(newStr);

        // Calculate missing values & verify math: Existing - Downgrade == New
        let isMathValid = true;
        let mathNote = "OK";

        if (existingNum !== null && downNum !== null) {
          const expected = existingNum - downNum;
          if (newNum !== null) {
            if (Math.abs(expected - newNum) > 0.1) {
              isMathValid = false;
              mathNote = `Mismatch: ${existingNum} - ${downNum} = ${expected} (Email says ${newNum})`;
            } else {
              mathNote = `Verified: ${existingNum} - ${downNum} = ${newNum} Mbps`;
            }
          } else {
            newNum = expected;
            mathNote = `Computed: ${existingNum} - ${downNum} = ${newNum} Mbps`;
          }
        }

        // Vendor inference
        let vendor = shortenVendor(vendorStr || doc.body.textContent);

        // Fallback distributor/BTS if not in cell
        if (!distriStr) {
          const guess = guessFromSubject(currentEmail.subject);
          distriStr = guess.distri || "";
        }

        if (linkId || existingNum !== null || downNum !== null) {
          const unitRate = calculateVendorRate(vendor, "D2D", newNum || existingNum || 0);

          results.push({
            included: true,
            linkId: linkId || "",
            distri: distriStr || "",
            vendor: vendor,
            existingBw: existingNum !== null ? String(existingNum) : "",
            downgradeBw: downNum !== null ? String(downNum) : "",
            newBw: newNum !== null ? String(newNum) : "",
            category: "D2D",
            unitCost: unitRate,
            effectiveDate: dateStr || currentEmail.date || "ASAP",
            mathNote: mathNote,
            isMathValid: isMathValid
          });
        }
      }
    }
  }

  return results;
}

/**
 * Text parsing fallback using regular expressions
 */
function parseTextBody(text) {
  const results = [];
  const linkMatches = text.match(/(?:lnk3_[A-Za-z0-9_]+|LT3[A-Za-z0-9_]+|\b\d{5,7}\b)/gi) || [];
  const uniqueLinks = Array.from(new Set(linkMatches));

  // Extract capacity numbers
  const existMatch = text.match(/(?:Existing|Current|Present)\s*(?:Capacity|BW|Bandwidth)[^\d]*?(\d+)/i);
  const downMatch = text.match(/(?:Downgrade|Down)\s*(?:Capacity|BW|Bandwidth)[^\d]*?(\d+)/i);
  const newMatch = text.match(/(?:New|Revised|After)\s*(?:Capacity|BW|Bandwidth)[^\d]*?(\d+)/i);

  const existNum = existMatch ? parseInt(existMatch[1]) : null;
  const downNum = downMatch ? parseInt(downMatch[1]) : null;
  let newNum = newMatch ? parseInt(newMatch[1]) : null;

  if (existNum !== null && downNum !== null && newNum === null) {
    newNum = existNum - downNum;
  }

  const guess = guessFromSubject(currentEmail.subject);
  const vendor = shortenVendor(text);

  if (uniqueLinks.length > 0) {
    uniqueLinks.forEach(link => {
      const unitRate = calculateVendorRate(vendor, "D2D", newNum || existNum || 0);
      results.push({
        included: true,
        linkId: link,
        distri: guess.distri || "",
        vendor: vendor,
        existingBw: existNum !== null ? String(existNum) : "",
        downgradeBw: downNum !== null ? String(downNum) : "",
        newBw: newNum !== null ? String(newNum) : "",
        category: "D2D",
        unitCost: unitRate,
        effectiveDate: currentEmail.date || "ASAP",
        mathNote: "Extracted via text patterns",
        isMathValid: true
      });
    });
  }

  return results;
}

/**
 * Extracts guess for BTS / Distributor and Link ID from Subject
 */
function guessFromSubject(subject) {
  const result = { distri: "", linkId: "", vendor: "SCL" };
  if (!subject) return result;

  // Search for link ID in subject
  const linkM = subject.match(/(lnk3_[A-Za-z0-9_]+)/i);
  if (linkM) result.linkId = linkM[1];

  // Search for BTS or Distributor pattern
  const btsM = subject.match(/(?:for|of|regarding)\s+([A-Za-z0-9\s\.\(\)\%\-]+?)(?:\s+link|\s+capacity|\s+downgrade|\s+bts|$)/i) ||
               subject.match(/([A-Za-z0-9\s\.\(\)\%\-]+(?:BTS|Distributor|Network|WiFi))/i);
  if (btsM) {
    result.distri = btsM[1].trim();
  }

  result.vendor = shortenVendor(subject);
  return result;
}

/**
 * Renders the editable cards in the taskpane
 */
function renderRecordsUI() {
  const container = document.getElementById("records-list");
  container.innerHTML = "";

  currentEmail.records.forEach((rec, idx) => {
    const card = document.createElement("div");
    card.className = `record-card ${rec.included ? 'included' : ''}`;
    card.id = `record-card-${idx}`;

    card.innerHTML = `
      <div class="record-card-header">
        <div class="record-title-group">
          <input type="checkbox" id="rec-inc-${idx}" ${rec.included ? 'checked' : ''} />
          <span class="record-title">Link #${idx + 1}: ${escapeHtml(rec.linkId || "New Link")}</span>
        </div>
        <button class="btn-danger-icon" id="rec-del-${idx}" title="Remove Link">🗑️</button>
      </div>

      <div class="record-grid">
        <div class="form-group-sm">
          <label class="form-label-sm">Provider (Col D)</label>
          <select id="rec-vendor-${idx}" class="form-control-sm">
            <option value="SCL" ${rec.vendor === 'SCL' ? 'selected' : ''}>SCL (Summit)</option>
            <option value="FHL" ${rec.vendor === 'FHL' ? 'selected' : ''}>FHL (Fiber@Home)</option>
            <option value="Bahon" ${rec.vendor === 'Bahon' ? 'selected' : ''}>Bahon</option>
            <option value="BTCL" ${rec.vendor === 'BTCL' ? 'selected' : ''}>BTCL</option>
            <option value="Earth" ${rec.vendor === 'Earth' ? 'selected' : ''}>Earth</option>
          </select>
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Link ID (Col E)</label>
          <input type="text" id="rec-link-${idx}" class="form-control-sm" value="${escapeHtml(rec.linkId)}" placeholder="e.g. lnk3_..." />
        </div>

        <div class="form-group-sm record-full-width">
          <label class="form-label-sm">BTS Name / Service Agent (Col B)</label>
          <input type="text" id="rec-distri-${idx}" class="form-control-sm" value="${escapeHtml(rec.distri)}" placeholder="e.g. Sunamganj BTS" />
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Existing BW (G)</label>
          <input type="number" id="rec-exist-${idx}" class="form-control-sm" value="${rec.existingBw}" placeholder="Mbps" />
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Downgrade BW (J)</label>
          <input type="number" id="rec-down-${idx}" class="form-control-sm" value="${rec.downgradeBw}" placeholder="Mbps" />
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">New BW (K = G - J)</label>
          <input type="number" id="rec-new-${idx}" class="form-control-sm" value="${rec.newBw}" placeholder="Mbps" />
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Slab Category</label>
          <select id="rec-cat-${idx}" class="form-control-sm">
            <option value="D2D" ${rec.category === 'D2D' ? 'selected' : ''}>D2D</option>
            <option value="LH" ${rec.category === 'LH' ? 'selected' : ''}>Long Haul (LH)</option>
            <option value="METRO" ${rec.category === 'METRO' ? 'selected' : ''}>Metro</option>
          </select>
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Unit Cost BDT (Col L)</label>
          <input type="text" id="rec-rate-${idx}" class="form-control-sm" value="${rec.unitCost}" />
        </div>

        <div class="form-group-sm">
          <label class="form-label-sm">Effective Date (Col N)</label>
          <input type="text" id="rec-date-${idx}" class="form-control-sm" value="${escapeHtml(rec.effectiveDate)}" />
        </div>
      </div>

      <div class="math-status ${rec.isMathValid ? 'ok' : 'mismatch'}" id="rec-math-${idx}">
        <span>${escapeHtml(rec.mathNote || "Math Verified")}</span>
        <span>${rec.isMathValid ? '✓' : '⚠️'}</span>
      </div>
    `;

    container.appendChild(card);
    bindRecordEvents(idx);
  });
}

function bindRecordEvents(idx) {
  const incCheck = document.getElementById(`rec-inc-${idx}`);
  const card = document.getElementById(`record-card-${idx}`);
  const linkInp = document.getElementById(`rec-link-${idx}`);
  const distriInp = document.getElementById(`rec-distri-${idx}`);
  const vendorSel = document.getElementById(`rec-vendor-${idx}`);
  const existInp = document.getElementById(`rec-exist-${idx}`);
  const downInp = document.getElementById(`rec-down-${idx}`);
  const newInp = document.getElementById(`rec-new-${idx}`);
  const catSel = document.getElementById(`rec-cat-${idx}`);
  const rateInp = document.getElementById(`rec-rate-${idx}`);
  const dateInp = document.getElementById(`rec-date-${idx}`);
  const delBtn = document.getElementById(`rec-del-${idx}`);

  // Inclusion toggle
  incCheck.addEventListener("change", (e) => {
    currentEmail.records[idx].included = e.target.checked;
    if (e.target.checked) card.classList.add("included");
    else card.classList.remove("included");
  });

  // Value updates
  linkInp.addEventListener("input", (e) => { currentEmail.records[idx].linkId = e.target.value.trim(); });
  distriInp.addEventListener("input", (e) => { currentEmail.records[idx].distri = e.target.value.trim(); });
  dateInp.addEventListener("input", (e) => { currentEmail.records[idx].effectiveDate = e.target.value.trim(); });

  // BW and Rate recalculations
  const updateBwAndRate = () => {
    const eVal = parseFloat(existInp.value) || 0;
    const dVal = parseFloat(downInp.value) || 0;
    const nVal = eVal - dVal;
    
    if (eVal > 0 && dVal > 0) {
      newInp.value = nVal >= 0 ? nVal : 0;
      currentEmail.records[idx].newBw = String(newInp.value);
    }
    
    currentEmail.records[idx].existingBw = existInp.value;
    currentEmail.records[idx].downgradeBw = downInp.value;

    // Recalculate rate
    const vendor = vendorSel.value;
    const cat = catSel.value;
    const bwToRate = parseFloat(newInp.value) || eVal;
    const calcRate = calculateVendorRate(vendor, cat, bwToRate);
    rateInp.value = calcRate;
    currentEmail.records[idx].unitCost = calcRate;
    currentEmail.records[idx].vendor = vendor;
    currentEmail.records[idx].category = cat;

    // Math check
    const mathBox = document.getElementById(`rec-math-${idx}`);
    if (mathBox) {
      if (eVal - dVal === parseFloat(newInp.value)) {
        mathBox.className = "math-status ok";
        mathBox.innerHTML = `<span>Verified: ${eVal} - ${dVal} = ${newInp.value} Mbps</span><span>✓</span>`;
        currentEmail.records[idx].isMathValid = true;
      } else {
        mathBox.className = "math-status mismatch";
        mathBox.innerHTML = `<span>Mismatch: ${eVal} - ${dVal} = ${eVal - dVal} != ${newInp.value}</span><span>⚠️</span>`;
        currentEmail.records[idx].isMathValid = false;
      }
    }
  };

  existInp.addEventListener("input", updateBwAndRate);
  downInp.addEventListener("input", updateBwAndRate);
  newInp.addEventListener("input", () => {
    currentEmail.records[idx].newBw = newInp.value;
    const calcRate = calculateVendorRate(vendorSel.value, catSel.value, parseFloat(newInp.value) || 0);
    rateInp.value = calcRate;
    currentEmail.records[idx].unitCost = calcRate;
  });
  vendorSel.addEventListener("change", updateBwAndRate);
  catSel.addEventListener("change", updateBwAndRate);
  rateInp.addEventListener("input", (e) => { currentEmail.records[idx].unitCost = e.target.value.trim(); });

  // Delete
  delBtn.addEventListener("click", () => {
    currentEmail.records.splice(idx, 1);
    renderRecordsUI();
    updateStatusBadge(`${currentEmail.records.length} Link(s)`, "info");
  });
}

function handleAddManualRecord() {
  currentEmail.records.push({
    included: true,
    linkId: "",
    distri: "",
    vendor: "SCL",
    existingBw: "",
    downgradeBw: "",
    newBw: "",
    category: "D2D",
    unitCost: "0.00",
    effectiveDate: currentEmail.date || "ASAP",
    mathNote: "Manually Added",
    isMathValid: true
  });
  renderRecordsUI();
}

/**
 * Sync selected rows to Power Automate Bridge (Destination: Downgrade Sheet)
 */
async function handleSyncClick() {
  const selectedRecords = currentEmail.records.filter(r => r.included && r.linkId);

  if (selectedRecords.length === 0) {
    showStatus("Please ensure at least one valid link (with Link ID) is checked.", "error");
    return;
  }

  const bridgeUrl = getActiveBridgeUrl();
  const syncBtn = document.getElementById("btn-sync-sheet");
  const syncIcon = document.getElementById("sync-icon");
  const syncText = document.getElementById("sync-btn-text");

  syncBtn.disabled = true;
  syncIcon.textContent = "⏳";
  syncText.textContent = `Syncing ${selectedRecords.length} record(s)...`;
  showStatus(`Submitting ${selectedRecords.length} record(s) to Power Automate -> Downgrade Sheet...`, "loading");

  let successCount = 0;
  let errorMsgs = [];

  for (let i = 0; i < selectedRecords.length; i++) {
    const rec = selectedRecords[i];
    
    // Exact schema matching the Downgrade Sheet + Universal Sync Bridge
    const payload = {
      sheet: "Downgrade",
      sheetName: "Downgrade",
      target: "Downgrade",
      date: currentEmail.date,
      month: new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }),
      distri: rec.distri || "N/A",
      btsName: rec.distri || "N/A",
      routeName: "",
      vendor: rec.vendor,
      provider: rec.vendor,
      linkId: rec.linkId,
      locationTagging: "",
      extCap: String(rec.existingBw || "0"),
      existingBw: rec.existingBw || "0",
      existingUnitCost: "",
      existingTotalMrc: "",
      downgradeBw: String(rec.downgradeBw || "0"),
      newCap: String(rec.newBw || "0"),
      newBw: rec.newBw || "0",
      unitPrice: rec.unitCost || "0.00",
      downgradedUnitCost: rec.unitCost || "0.00",
      downgradedTotalMrc: ((parseFloat(rec.newBw) || 0) * (parseFloat(rec.unitCost) || 0)).toFixed(2),
      effectiveDate: rec.effectiveDate || "ASAP",
      activeCustomers: "0",
      initialId: `Link: ${rec.linkId}`,
      reqNo: (currentEmail.subject.match(/(?:REQ-|L3R-[A-Z]+-\d+-\d+-REQ-)(\d+)/i) || ["","UDD-DOWNGRADE"])[1],
      item: "NTTN Capacity Downgrade",
      zone: "N/A",

      // Literal Excel Table Column Headers (Case-insensitive matching by header name)
      "Date": currentEmail.date,
      "BTS Name/Service Agent": rec.distri || "",
      "Route Name": "",
      "Provider": rec.vendor,
      "Link ID": rec.linkId,
      "Location Tagging": "",
      "Existing BW": rec.existingBw || "0",
      "Existing Unit Cost": "",
      "Existing Total MRC": "",
      "Downgrade BW (Mbps)": rec.downgradeBw || "0",
      "New BW after Down-gradation (Mbps)": rec.newBw || "0",
      "Downgraded Unit Cost": rec.unitCost || "0.00",
      "Downgraded Total MRC": ((parseFloat(rec.newBw) || 0) * (parseFloat(rec.unitCost) || 0)).toFixed(2),
      "Effective Date": rec.effectiveDate || "ASAP"
    };

    try {
      console.log(`[SYNC] Sending record ${i + 1}/${selectedRecords.length}:`, payload);
      
      const response = await fetch(bridgeUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });

      if (response.ok || response.status === 200 || response.status === 202) {
        successCount++;
      } else {
        const errorText = await response.text();
        errorMsgs.push(`Link ${rec.linkId}: HTTP ${response.status} - ${errorText.substring(0, 100)}`);
      }
    } catch (err) {
      console.error(`[SYNC ERROR] Link ${rec.linkId}:`, err);
      errorMsgs.push(`Link ${rec.linkId}: ${err.message}`);
    }
  }

  syncBtn.disabled = false;
  if (successCount === selectedRecords.length) {
    syncIcon.textContent = "✅";
    syncText.textContent = "Synced to Downgrade Sheet ✓";
    showStatus(`Successfully synced all ${successCount} record(s) to the Downgrade Sheet!`, "success");
    setTimeout(() => {
      syncIcon.textContent = "🔄";
      syncText.textContent = "Sync to Downgrade Sheet";
    }, 4000);
  } else if (successCount > 0) {
    syncIcon.textContent = "⚠️";
    syncText.textContent = `Partial (${successCount}/${selectedRecords.length})`;
    showStatus(`Synced ${successCount} record(s), but encountered errors:\n${errorMsgs.join("\n")}`, "error");
  } else {
    syncIcon.textContent = "❌";
    syncText.textContent = "Sync Failed";
    showStatus(`Failed to sync to Power Automate:\n${errorMsgs.join("\n")}`, "error");
  }
}

/**
 * Copies selected rows formatted for direct Excel Ctrl+V at cell A{row}
 * Columns A through N:
 * A: Date, B: BTS Formula, C: Route, D: Provider, E: Link ID, F: Loc, G: Ext BW,
 * H: Ext Cost, I: Ext MRC, J: Down BW, K: New BW Formula, L: Down Cost, M: Down MRC, N: Effective Date
 */
function handleCopyTsvClick() {
  const selectedRecords = currentEmail.records.filter(r => r.included && r.linkId);

  if (selectedRecords.length === 0) {
    showStatus("Please ensure at least one valid link is checked to copy.", "error");
    return;
  }

  const lines = selectedRecords.map(r => {
    return [
      r.effectiveDate || currentEmail.date,          // Col A: Date
      r.distri ? r.distri : '=XLOOKUP(E{row},Sheet1!C:C,Sheet1!O:O)', // Col B: BTS Name
      "",                                            // Col C: Route Name
      r.vendor,                                      // Col D: Provider
      r.linkId,                                      // Col E: Link ID
      "",                                            // Col F: Location Tagging
      r.existingBw || "0",                           // Col G: Existing BW
      "",                                            // Col H: Existing Unit Cost
      "",                                            // Col I: Existing Total MRC
      r.downgradeBw || "0",                          // Col J: Downgrade BW
      `=G{row}-J{row}`,                              // Col K: New BW Formula
      r.unitCost || "",                              // Col L: Downgraded Unit Cost
      "",                                            // Col M: Downgraded Total MRC
      r.effectiveDate || "ASAP"                      // Col N: Effective Date
    ].join("\t");
  });

  const tsvText = lines.join("\n");

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(tsvText).then(() => {
      showStatus(`Copied ${selectedRecords.length} row(s) to clipboard! Paste directly into cell A in the Downgrade sheet.`, "success");
    }).catch(err => {
      fallbackCopy(tsvText);
    });
  } else {
    fallbackCopy(tsvText);
  }
}

function fallbackCopy(text) {
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.style.position = "fixed";
  ta.style.left = "-9999px";
  document.body.appendChild(ta);
  ta.select();
  try {
    document.execCommand("copy");
    showStatus(`Copied to clipboard! Paste directly into cell A in the Downgrade sheet.`, "success");
  } catch (err) {
    showStatus("Could not copy automatically. Please select text manually.", "error");
  }
  document.body.removeChild(ta);
}

function updateStatusBadge(text, type) {
  const badge = document.getElementById("detection-status-badge");
  if (badge) {
    badge.textContent = text;
    badge.className = `badge ${type || ''}`;
  }
}

function showStatus(msg, type) {
  const box = document.getElementById("status-box");
  const msgEl = document.getElementById("status-message");
  const iconEl = document.getElementById("status-icon");

  box.style.display = "flex";
  box.className = `status-box ${type}`;
  msgEl.innerText = msg;

  if (type === "success") iconEl.textContent = "✅";
  else if (type === "error") iconEl.textContent = "❌";
  else if (type === "loading") iconEl.textContent = "⏳";
  else iconEl.textContent = "ℹ️";
}

function hideStatus() {
  const box = document.getElementById("status-box");
  if (box) box.style.display = "none";
}

function extractNumeric(val) {
  if (!val) return null;
  const cleaned = String(val).replace(/[^0-9\.]/g, "");
  const num = parseFloat(cleaned);
  return isNaN(num) ? null : num;
}

function stripHtml(html) {
  const tmp = document.createElement("div");
  tmp.innerHTML = html;
  return tmp.textContent || tmp.innerText || "";
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
