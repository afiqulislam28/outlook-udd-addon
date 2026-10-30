# Outlook Add-in: Complete Deployment & User Guide
**Supports: Outlook Android, Outlook Desktop (Windows / Mac), and Outlook on the Web (Microsoft 365)**

---

## 1. How Outlook Add-ins Work Across Devices

Microsoft Office Add-ins are **cloud-synced web applications**. 
- The **`manifest.xml`** file tells Microsoft Outlook where your button should appear and where your web files (`taskpane.html`, `taskpane.js`) are hosted.
- **Cross-Platform Sync**: When you install or sideload an add-in into your Microsoft 365 / Outlook account, **Microsoft Exchange automatically syncs the add-in to all your connected devices**, including:
  1. **Outlook on the Web (Microsoft 365)**
  2. **Outlook Desktop** (New Outlook & Classic Outlook for Windows / Mac)
  3. **Outlook Mobile App on Android** (and iOS)

---

## 2. Where the Button Appears on Each Platform

When you open any email (like the *Capacity Downgrade* email):

| Platform | Where to find the "Assign Task" button |
| :--- | :--- |
| **Outlook on the Web (365)** | In the email action bar (next to Reply / Forward) or under the **`...` (More actions)** menu on the email. |
| **Outlook Desktop** | On the top Ribbon under the **Task Actions** group (`Assign Task`). In New Outlook, also in the message action bar. |
| **Outlook Android** | At the bottom action bar when viewing the email, or tap the **`...` (Three dots)** inside the open email header/footer. |

---

## 3. Step 1: Host Your Web Files (HTTPS Required)

Microsoft Outlook requires all add-in URLs (`https://...`) to be served over **HTTPS with a valid SSL certificate**.

You have several easy options:

### Option A: Free Hosting on GitHub Pages (Recommended for Quick Setup)
1. Push this folder to a GitHub repository.
2. In GitHub Repo Settings -> **Pages**, select `main` branch `/ (root)` and click **Save**.
3. You will get a free HTTPS URL like: `https://your-username.github.io/outlook-addon/`
4. Replace `https://localhost:3000` inside `manifest.xml` with your GitHub Pages URL!

### Option B: Local Testing with Ngrok (for development)
1. Start any local web server on port 3000 (e.g. using VS Code Live Server or Python: `python -m http.server 3000`).
2. Run ngrok:
   ```bash
   ngrok http 3000
   ```
3. Copy the `https://xxxx.ngrok-free.app` URL and update `manifest.xml`.

---

## 4. Step 2: Sideload & Install the Add-in

Once your files are hosted (or via local HTTPS), sideload it into your Outlook account:

### Method 1: Sideload via Outlook on the Web (Fastest for testing)
1. Open **[Outlook on the Web (outlook.office.com)](https://outlook.office.com)** in your browser and sign in to your Microsoft 365 work/school account.
2. Open any email.
3. Click the **`...` (More actions)** icon at the top-right of the email.
4. Select **Get Add-ins** (or **Apps** -> **Add apps**).
5. In the dialog that opens, select **My add-ins** on the left menu.
6. Scroll down to **Custom Add-ins**, click **Add a custom add-in** -> **Add from file...**
7. Select your **`manifest.xml`** file.
8. Click **Install / Warning Proceed**.

> **Note for Android:** Because mobile apps sync directly with your Microsoft 365 cloud mailbox, once installed via Outlook on the Web, **close and reopen the Outlook app on your Android phone**. The "Assign Task" button will appear on your emails!

---

### Method 2: Organization-Wide Deployment (Microsoft 365 Admin Center)
If your organization admin wants to deploy this button to you and your teammates (like the Sales & Marketing / Transmission teams in your screenshot):
1. Go to **[admin.microsoft.com](https://admin.microsoft.com)**.
2. Navigate to **Settings** -> **Integrated apps**.
3. Click **Upload custom apps**.
4. Choose **Office Add-in** and upload `manifest.xml`.
5. Choose which users/groups receive the add-in (e.g., Specific users, S&M Team, Transmission Team, or Entire organization).
6. Click **Deploy**. It will automatically appear on all their Outlook apps (Desktop, Web, Android).

---

## 5. Step 3: Assigning Your Custom Task / Workflow

In `src/taskpane.js`, find the function:

```javascript
async function executeAssignedTask(taskPayload) {
  // taskPayload includes:
  // - taskPayload.taskTitle
  // - taskPayload.assignee ("Transmission Team", "S&M Team", etc.)
  // - taskPayload.priority ("High", "Normal", "Low")
  // - taskPayload.notes
  // - taskPayload.email.subject ("Capacity Downgrade")
  // - taskPayload.email.from (Sender name & email)
  // - taskPayload.extractedData (Link ID "lnk3_010925_075_nb", Current Capacity, Downgrade Capacity, New Capacity)
  
  // Connect your API here:
  const response = await fetch("https://your-api-domain.com/tasks", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(taskPayload)
  });
  return await response.json();
}
```

You can point this to:
- A custom internal REST API or backend server
- A **Microsoft Power Automate** HTTP webhook flow
- A **Jira / ServiceNow / Trello** task creation API
- A **Microsoft Teams / Slack** incoming webhook
