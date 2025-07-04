# 🛠️ Windows Update Compliance Tracker

This project provides a complete solution for **Windows Update Compliance Tracking** using a combination of:

- **Google Apps Script** (as a Web App)
- **PowerShell** (run on client machines)

The system collects and visualizes update compliance data across endpoints, enabling proactive compliance management through dashboards and notifications.

---

## 📦 Project Components

### 1. **Google Apps Script Web App**
A Google Apps Script deployed as a Web App that:
- Receives JSON POST requests from the PowerShell script
- Logs update data in a Google Sheet
- Sends email alerts for non-compliance
- Provides a live dashboard and export tools

### 2. **PowerShell Update Script**
A client-side PowerShell script that:
- Gathers system and update info
- Installs updates (if available)
- Prompts the user interactively for reboot
- Sends a compliance report to the webhook URL

---

## 🚀 Setup Instructions

### 🔧 Google Apps Script (Backend)

1. **Create Google Sheet**
   - Name the sheet: `Windows Update Compliance`.

2. **Open the Apps Script Editor**
   - Extensions → Apps Script.

3. **Paste the Provided Google Script**
   - Use the full code block provided in `googlesheets.gs`.

4. **Deploy as Web App**
   - Click `Deploy > Manage Deployments`.
   - Set access to "Anyone with the link" or per your security policy.
   - Copy the **Web App URL**.

5. **Set Email Recipients**
   - Update `COMPLIANCE_EMAIL` and `NOTIFICATION_EMAIL` in the script.

---

### ⚙️ PowerShell Script (Client)

1. **Install `PSWindowsUpdate` Module** (if not present)
2. **Update Webhook URL**
   - Set `$GoogleSheetsWebhookUrl` at the top of the script.

3. **Schedule Execution**
   - Deploy as a Scheduled Task with elevated privileges.

4. **Interactive Reboot**
   - Uses a Scheduled Task to prompt the active user.
   - Logs the user's choice or timeout.

---

## 📊 Features

### ✅ Google Apps Script

- **Data Logging**: Captures update status, OS info, user data.
- **Email Alerts**:
  - Non-compliant systems
  - Reboot postponed
  - Daily summary (5–6 PM)
- **Live Dashboard**:
  - Compliance rates
  - 7‑day trends
  - Non-compliant system listing
- **Export Tool**: Create CSV of non-compliant systems.
- **Data Cleanup**: Monthly purge of records older than 30 days.

### 💻 PowerShell Script

- **Scans for Updates**
- **Installs Patches**
- **Detects Reboot Requirement**
- **Interactive User Prompt for Reboot**
- **Sends Report to Google Sheets**

---

## 📐 Data Format

| Field                     | Description                                |
|--------------------------|--------------------------------------------|
| `Timestamp`              | Script run time in ISO format              |
| `ComputerName`, `UserName`, `CurrentUser` | Machine & user identity         |
| `OSVersion`, `OSBuild`   | Windows version details                    |
| `AvailableUpdates`       | Number of updates detected                 |
| `UpdatesInstalled`       | Number of updates installed                |
| `RebootRequired`         | Boolean indicating reboot need             |
| `RebootAction`           | Immediate, postponed, or skipped           |
| `ComplianceStatus`       | Final compliance state                     |
| `IP Address`, `Domain`, `RAM`, etc. | Additional system info           |
| `ExecutionDuration`      | Runtime duration in minutes                |
| `InstallDetails`         | JSON array of updates and statuses         |

---

## 📢 Notifications

- **Non-Compliant Alert** → `COMPLIANCE_EMAIL`
- **Reboot Postponed** → `NOTIFICATION_EMAIL`
- **Daily Summary** (auto-sent between 5–6 PM)

---

## 🧪 Testing

1. Run the PowerShell script manually.
2. Check the `%TEMP%` log file.
3. Inspect Google Sheet and Dashboard.
4. Verify email notifications are triggered.

---

## 🔐 Security

- Restrict Web App access to authorized clients.
- Align with your organization’s security and privacy policies.

---

## 📎 Authors

- **Arun Jose** – Compliance Automation Engineer (📧 `your_email`)  
- **Script Maintainer** – (📧 `another_email`)

---

## 🏁 Version History

- **v3.1** – “*It Just Works*” Edition  
  - Added reliable reboot via `shutdown.exe`  
  - Enhanced dashboard reporting  
  - Robust error handling and modular structure  

---

## 📝 Appendix

Admin menu items in the Google Sheet:
- **Compliance Tools → Refresh Dashboard**
- **Send Summary Email**
- **Export Non-Compliant Systems**

---

## ✅ Download
download as zip maybe?
```
