# ShowOps v3.0 — Team Operations Dashboard

A single-page operations dashboard for managing shows, tasks, blockers, and team workload.  
Works fully offline (localStorage) or live with Google Sheets as the data store.

## Architecture

```
GitHub Pages (index.html)
        │
        ▼
Google Apps Script Web App (Code.gs)
        │
        ▼
Google Sheets (data store)
```

The dashboard polls the backend every 10 seconds. Sheet-side edits appear in the UI automatically.  
All dashboard write actions (raise request, assign task, update status, raise blocker, etc.) are  
written to both localStorage **and** Google Sheets when `useBackend: true`.

---

## Setup

### 1. Google Sheet

Open the sheet: https://docs.google.com/spreadsheets/d/1fHV_8KDZjUD1B2cl4eHRl92Hl0DgVxoB3fj4RPWH_Tg

Go to **Extensions → Apps Script**, paste `Code.gs`, then run `setupSheets()` **once** to create all tabs and headers.

### 2. Deploy Apps Script

**Deploy → New deployment → Web app**  
- Execute as: **Me**  
- Who has access: **Anyone** (or "Anyone in your domain" for Workspace)  

Copy the web app URL.

### 3. Connect the dashboard

In `index.html`, update `SHOWOPS_CONFIG`:

```js
window.SHOWOPS_CONFIG = {
  apiUrl: "YOUR_APPS_SCRIPT_WEB_APP_URL",
  useBackend: true,
  pollMs: 10000
};
```

Push to GitHub and enable **GitHub Pages** (Settings → Pages → Branch: main → / root).

---

## Google Sheet tabs

| Tab | Key columns |
|-----|-------------|
| `Shows_Master` | Show_ID, Show_Name, Show_Type, Genre, Language, Active |
| `Tasks` | Task_ID, Show_ID, Task, Raised_By, Assigned_To, Status, HML, P_Level, High_Priority, Deadline, Created_At, Updated_At, Latest_Message |
| `Issues` | Issue_ID, Show_ID, Task_ID, Issue, Severity, Status, Raised_By, Owner, Created_At |
| `People_Access` | User_ID, Name, Email, Role, After_Hours, Active |
| `Activity_Log` | Date_Time, User, Action, Show, Details |
| `Settings` | Key, Value |
| `Task_Updates` | Task_ID, Status, Message, Updated_By, Updated_At |
| `Requests` | Request_ID, Show_ID, Task, Raised_By, … |
| `Weekly_Performance` | Week, Assigned, Completed, Blocked, Completion_Rate, Avg_TAT |

---

## What syncs to Sheets

Every dashboard action writes to Sheets when backend is enabled:

- Raise request / assign task / bulk reassign  
- Update task status (In Progress → Completed → Reopen)  
- Raise / update blocker  
- Add / archive show  
- Add / deactivate person  
- Change settings (request window, after-hours access)  
- All actions append to `Activity_Log`

---

## Rules

- Show Type mandatory: UK / EU / Other  
- Request window: 06:00–18:30 (configurable in Settings)  
- Selected after-hours users can raise requests outside the window  
- H/M/L + P0/P1/P2 + High Priority flag  
- My Tasks sorted: High flag → P0/P1/P2 → H/M/L → deadline  
- Completed tasks move to the bottom; can be reopened  
- Show status derived from latest task + open blockers  
- People & Access, Performance, Settings, Show archival: Manager only  

---

## Security notes

- Do **not** commit API keys, service-account keys, or Google credentials  
- Enforce permissions server-side in Apps Script  
- Use Google Workspace domain access where possible so `Session.getActiveUser()` identifies the signed-in user
