# UDD Downgrade Sync — Outlook Add-in
**Direct Cloud Sync from Outlook (Desktop, Web & Android) to "DISTRIBUTOR REQUEST - AGM PLANNING.xlsx" (Sheet: Downgrade)**

---

## 1. What This Add-in Does

When you open a **Capacity Downgrade** email from the Transmission Team:
1. **Auto-Extracts**:
   - Link ID (`lnk3_...`, `LT3...`, etc.)
   - Provider / Vendor (SCL, FHL, Bahon, BTCL, Earth)
   - BTS Name / Service Agent / Distributor
   - Existing Bandwidth (Col G)
   - Downgrade Bandwidth (Col J)
   - New Bandwidth (Col K = Existing - Downgrade)
   - SCL / FHL Slab Rate Calculation with 5% VAT (Col L)
   - Effective Date (Col N)
2. **Validates Math**:
   - Checks that $\text{Existing} - \text{Downgrade} == \text{New}$.
   - Displays a green checkmark or yellow warning if there's any discrepancy.
3. **Dual Destination Options**:
   - **🔄 Sync to Downgrade Sheet**: Dispatches HTTP POST to the Power Automate Flow endpoint (`BRIDGE_URL`), which appends the row directly into the **Downgrade** worksheet in the cloud Excel file.
   - **📋 Copy for Excel (Ctrl+V)**: Copies clean Tab-Separated Values (TSV) formatted for columns A through N, allowing instant paste at cell `A{next_row}` if working directly inside the Excel desktop app.

---

## 2. Power Automate Bridge Connection

The add-in uses your Power Automate HTTP trigger:
```text
https://defaultbda5f9dab8cd47efba3f91645269a2.48.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/2ab50687df5a4a978c31bf38d705121c/triggers/manual/paths/invoke?...
```

### JSON Payload sent for each Downgrade Link:
```json
{
  "sheet": "Downgrade",
  "sheetName": "Downgrade",
  "target": "Downgrade",
  "date": "30-Oct-26",
  "month": "30 Oct 2026",
  "distri": "Sunamganj BTS",
  "btsName": "Sunamganj BTS",
  "routeName": "",
  "vendor": "SCL",
  "provider": "SCL",
  "linkId": "lnk3_291021_003_nb",
  "locationTagging": "",
  "extCap": "1950",
  "existingBw": "1950",
  "existingUnitCost": "",
  "existingTotalMrc": "",
  "downgradeBw": "150",
  "newCap": "1800",
  "newBw": "1800",
  "unitPrice": "15.75",
  "downgradedUnitCost": "15.75",
  "downgradedTotalMrc": "28350.00",
  "effectiveDate": "ASAP",
  "activeCustomers": "0",
  "initialId": "Link: lnk3_291021_003_nb",
  "reqNo": "UDD-DOWNGRADE",
  "item": "NTTN Capacity Downgrade",
  "zone": "N/A"
}
```

---

## 3. How to Install / Sideload the Add-in

### Step 1: Host Web Files (HTTPS Required)
Microsoft Outlook requires HTTPS. You can:
- **Option A (GitHub Pages)**:
  Host this folder on GitHub Pages (e.g. `https://afiqulislam28.github.io/outlook-udd-addon/`).
  The URLs in `manifest.xml` are already configured for this address.
- **Option B (Local Dev / Ngrok)**:
  Run a local server with `python -m http.server 3000` and tunnel via `ngrok http 3000`. Replace URLs in `manifest.xml` with your ngrok URL.

### Step 2: Sideload into Outlook
1. Open **[Outlook on the Web (outlook.office.com)](https://outlook.office.com)**.
2. Open any email.
3. Click the **`...` (More actions)** icon on the email header.
4. Select **Get Add-ins** (or **Apps** -> **Add apps**).
5. Select **My add-ins** on the left menu.
6. Scroll down to **Custom Add-ins**, click **Add a custom add-in** -> **Add from file...**.
7. Select **`manifest.xml`**.
8. Click **Install / Proceed**.

> **Automatic Cross-Platform Sync:**
> Once sideloaded in Outlook Web, Microsoft Exchange automatically syncs the button to your:
> - **Outlook Desktop (Windows / Mac)**
> - **Outlook Android Mobile App**
