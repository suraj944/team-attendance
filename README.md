# Team Attendance

A daily Login / Logout website for calling teams. Callers sign in, post their Login and Logout once a day, and the message opens in WhatsApp in the team's format so they can send it to their group. Admins manage callers and branches and see attendance reports. All data is stored in a Google Sheet.

## Features

**Caller**
- Sign in with a Login ID and password given by the admin
- One Login and one Logout per day, with a warning on a second attempt
- Login time chosen with 12-hour dropdowns (no future time); Logout time is taken from the server and cannot be edited
- Date, caller name and SM are locked; the Logout SM is taken from that day's Login
- Confirmation step ("This cannot be changed later") before posting
- WhatsApp opens with the message ready, so the caller chooses the group
- "Kindly post the logout" reminder after the reminder time, and a warning for missed logouts on earlier days

**Admin**
- Today: who logged in, logged out or is pending, with a button that sends a logout reminder to a WhatsApp group
- Reports: any date range, branch summary, employee attendance, late logins, missed logouts, calls, connect %, leads, site visits, bookings, top performers and a day-by-day view per caller, with Excel download
- Callers: add with first name, last name, +91 phone, branch and joining date. The Login ID (first name + last name + last 4 phone digits) and password are created automatically and can be sent to the caller on WhatsApp
- Edit callers, reset passwords, mark callers inactive, and change branches with a history
- Settings: branches, SM names, locations, logout reminder time, late-login time and weekly off (default Monday)

## Files

| Path | What it is |
|---|---|
| `apps-script/Code.gs` | Server code: Google Sheet database, sign-in, rules, reports |
| `apps-script/Index.html` | The website (caller screen and admin panel) |
| `apps-script/appsscript.json` | Apps Script settings (India time zone, web app access) |
| `demo/index.html` | Test version that runs in a browser with sample data. It saves only in that browser and is not connected to Google Sheets |

## Google Sheet structure

`setup` creates these tabs:

| Tab | Contents |
|---|---|
| Employees | One row per employee: Employee ID (Login ID), first and last name, full name, phone, branch, joining date, role, status, created on, password hash (hidden) |
| Attendance | One row per employee per day: login and logout side by side, late flag, calls, connect %, leads, site visits, booking, status (Complete / Logout pending / Logout missed), admin note |
| Employee Summary | Monthly figures per employee, calculated by formulas. Change the month in B1 |
| Branch Summary | Monthly figures per branch, calculated by formulas. Change the month in B1 |
| Branch Transfers | History of branch changes |
| Branches, SMs, Locations | Dropdown lists used by the website |
| Settings | Logout reminder time, late-after time, weekly off |

A nightly job at 11 PM marks any day still "Logout pending" as "Logout missed".

## Setup

1. Open the Google Sheet (or create a blank one) and go to **Extensions > Apps Script**.
2. Replace the contents of `Code.gs` with `apps-script/Code.gs`.
3. Add an HTML file named exactly `Index` and paste `apps-script/Index.html` into it. Save.
4. Choose the function **setup** and press **Run**. Allow the permissions when asked. If Google shows "Google hasn't verified this app", choose **Advanced > Go to … (unsafe) > Allow**; this is your own script.
5. **Deploy > New deployment > Web app**. Execute as: **Me**. Who has access: **Anyone**. Copy the Web app URL and share it with the team.
6. Sign in as `admin` with password `admin123`, then open **Callers > Edit** on Admin and press **Reset password**.

### Updating the code later

Paste the new code, then **Deploy > Manage deployments > Edit > Version: New version > Deploy**. The website link stays the same.

### Optional: push with clasp

```bash
npm install -g @google/clasp
clasp login
cd apps-script
clasp clone <SCRIPT_ID>   # or: clasp create --type sheets --title "Team Attendance"
clasp push
```

## Notes

- Add callers, change names and reset passwords from the admin panel rather than typing in the sheet, so Login IDs and passwords are created correctly.
- Passwords are stored as SHA-256 hashes. Sign-in sessions last until midnight.
- A website cannot post into a WhatsApp group on its own; the caller (or admin, for reminders) picks the group and presses send.
