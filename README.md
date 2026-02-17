# 📅 Client Schedule Sync — README
### Excel → Outlook Calendar Sync Tool
---

## What Does This Tool Do?

This tool reads a list of clients and their important deadlines from your Excel spreadsheet and automatically creates calendar reminders in Microsoft Outlook for each one.

Think of it like setting an alarm clock — but instead of one alarm, it can set **multiple alarms leading up to a deadline** so you never get caught off guard.

When you run it again later, it **cleans up the old reminders first** and then creates fresh ones, so your calendar never gets cluttered with outdated entries.

---

## ✅ Compatibility

This tool works with:
- **Microsoft 365** (formerly Office 365) — fully supported, recommended
- **Microsoft Office 2016, 2019, 2021** — fully supported
- **Windows only** — this tool does not work on Mac because it uses Windows-specific Outlook automation

> 💡 If you use Microsoft 365, your calendar events will automatically sync across all your devices (desktop, phone, and tablet) because they live in your Exchange/cloud calendar.

---

## 🔧 One-Time Setup (Do This Before First Use)

You only need to do these steps **once**. After that the tool will just work every time you open the file.

---

### Step 1 — Enable Macros in Excel

This tool runs as a **macro** (a small program built into the Excel file). Excel blocks macros by default for security, so you need to allow them first.

**How to enable macros:**

1. Open Excel
2. Click **File** in the top left
3. Click **Options** at the very bottom of the left menu
4. Click **Trust Center** in the window that opens
5. Click the **Trust Center Settings...** button on the right
6. Click **Macro Settings** in the left panel
7. Select **"Disable all macros with notification"**
8. Click **OK**, then **OK** again
9. **Close and reopen the Excel file**

When you reopen the file, a yellow bar will appear at the top saying *"Macros have been disabled."* Click **Enable Content** on that bar.

> ⚠️ If you skip this step, clicking the Sync button will appear to do nothing at all.

---

### Step 2 — Enable the Required VBA References

The tool uses two Microsoft libraries that need to be switched on inside Excel. These are called **VBA References** and they tell Excel which extra features the tool is allowed to use.

**How to enable them:**

1. With the Excel file open, press **Alt + F11** on your keyboard — this opens the Visual Basic Editor (a separate window)
2. In the menu bar at the top of that window, click **Tools**
3. Click **References...**
4. A list will appear. Scroll through it and look for these two items. Put a **checkmark** next to each one by clicking the box beside them:

| Reference to enable | What it's for |
|---------------------|--------------|
| **Microsoft Outlook 16.0 Object Library** | Lets Excel talk to Outlook to create calendar events |
| **Microsoft Scripting Runtime** | Used internally to count and sort your alert types for the summary message |

> 💡 The number in "Outlook 16.0" might be slightly different on your computer (like 15.0 or 16.0) — that is fine, just check whichever Outlook version appears in the list.

5. Click **OK**
6. Close the Visual Basic Editor window (click the X or press **Alt + F4**)

> ⚠️ If these references are not enabled, you will see an error like *"User-defined type not defined"* or *"Can't create object"* when you try to run the sync. Enabling them fixes it.

---

### Step 3 — Make Sure Outlook Is Set Up

The tool will launch Outlook automatically if it is not already open, but Outlook must be **installed and signed in** on your computer with your email account before running the sync.

To check:
1. Open Outlook manually at least once
2. Make sure your email account is connected and your calendar is visible
3. Close Outlook (the tool will reopen it when needed)

---

## 📁 File Setup

The Excel file should have the following sheets:

| Sheet name | Purpose |
|------------|---------|
| Your main client sheet | Where client rows and deadlines live — this should be the **active sheet** when you click the Sync button |
| **Holidays** (optional) | List of holiday dates in column A starting from row A2 — the tool uses this to avoid placing reminders on non-working days |

**Your data must start on row 5.** Rows 1–4 are assumed to be headers or title rows and are skipped automatically.

---

## Before You Run It Each Time

- ✅ Excel file is open and you are on the client schedule sheet
- ✅ Microsoft Outlook is installed and signed in
- ✅ You have clicked **Enable Content** on the yellow macro bar (first time only)
- ✅ VBA References are enabled (first time only — see Step 2 above)

---

## How to Run It

1. Open your Excel file
2. Make sure you are on the **client schedule sheet**
3. Click the **"Sync Client Schedules"** button
4. A status bar at the bottom of Excel will show progress (e.g. *"Syncing Row 5 of 12..."*)
5. A summary message will appear when complete

---

## Understanding the Spreadsheet Columns

---

### Column A — Enabled (Yes / No)

This is the **master on/off switch** for each client row.

- Type **Yes** → the tool WILL create calendar reminders for this client
- Type **No** → the tool will SKIP this client entirely, no matter what else is filled in

> 💡 Think of it like a light switch. If the switch is off, nothing else in that row matters.

---

### Column B — Client Name

The name of the client. This shows up in the calendar event title and body so you know who the reminder is for.

---

### Column C — Employee Count

The number of employees at that client. This is written into the calendar event body for your reference.

---

### Column D — Insurance Agent Email

The email address of the insurance agent assigned to this client. They will be added as a **required attendee** on the calendar event, meaning Outlook will send them a meeting invite.

---

### Column E — Contact Email

The HR or main contact email at the client's company. This is written into the event body for easy reference.

---

### Column F — Contact Phone

The client's phone number. The tool automatically formats this as a **tappable phone link** so that when you open the calendar event on your phone in Outlook mobile, you can tap the number to dial directly — no typing needed.

You can enter the number in any format you like:
- `(813) 601-9216`
- `813-601-9216`
- `8136019216`

The tool will clean it up automatically.

---

### Columns G, L, Q — Scheduled (Yes / No)

Each client can have **up to 3 separate events** (for example: a 1095-C filing deadline, a Form 5500 deadline, and a PCORI Fee deadline).

These three columns are the **on/off switches for each individual event** within a client row.

- Column **G** controls Event 1
- Column **L** controls Event 2
- Column **Q** controls Event 3

> 💡 You can have the master switch (Column A) set to Yes but still turn off individual events here. For example: a client is active, but they don't need the Form 5500 reminder this year — just set that event's Scheduled column to No.

---

### Columns H, M, R — Event Type

The short name or label for the event (e.g. `1095-C`, `Form 5500`, `PCORI Fee`). This appears in the calendar event title.

---

### Columns I, N, S — Description

A longer description of what the event is about (e.g. `ALE Filing`, `Large Group Audit`). This is written into the event body.

---

### Columns J, O, T — Alerts

This is where you tell the tool **how many advance warning reminders** to create before the deadline.

**Format:** `90d:Y | 30d:Y | 7d:Y`

Each piece is called an **alert token** and follows this simple pattern:

```
[number of days before deadline]d:[Y or N]
```

| Token | What it means |
|-------|--------------|
| `90d:Y` | Create a reminder 90 days before the deadline — **Y = Yes, create it** |
| `30d:Y` | Create a reminder 30 days before the deadline — **Y = Yes, create it** |
| `7d:Y`  | Create a reminder 7 days before the deadline — **Y = Yes, create it** |
| `7d:N`  | Do NOT create a 7-day reminder — **N = No, skip it** |

Separate each token with a pipe character ` | `

**Example:**
```
90d:Y | 30d:Y | 7d:Y
```
This creates three advance alerts: 90 days out, 30 days out, and 7 days out — **plus** the deadline event itself. That is **4 calendar entries total** for that one event.

> 💡 Think of it like setting multiple alarm clocks for the same appointment. One fires 90 days before, one 30 days before, one 7 days before, and the final one on the deadline day itself.

**To disable a specific alert without removing it**, change its `:Y` to `:N`:
```
90d:Y | 30d:N | 7d:Y
```
This creates the 90-day and 7-day alerts but skips the 30-day one.

> 💡 Spaces around the `|` are fine — the tool handles them automatically.

---

### Columns K, P, U — Deadline

The actual due date for the event (e.g. `3/2/2026`).

The tool will:
1. Create a **DEADLINE** calendar event on this date (automatically moved to the prior Friday if it falls on a weekend or holiday)
2. Count backwards from this date to place all your advance alert reminders

> 💡 The deadline is the finish line. All your alerts count backwards from it.

---

## The Holidays Sheet

If you have a sheet named exactly **Holidays** with dates in column A starting from row 2, the tool will automatically **avoid those dates** when placing reminders.

For example, if a 7-day alert would land on Thanksgiving, the tool moves it to the Wednesday before instead.

This keeps all your reminders on actual working days.

---

## What Gets Created in Outlook

For each enabled event the tool creates calendar entries titled like this:

```
DEADLINE: 1095-C | Globex
-7d Alert: 1095-C | Globex
-30d Alert: 1095-C | Globex
-90d Alert: 1095-C | Globex
```

Inside each event you will see:

```
Client: Globex
─────────────────────────
Event: 1095-C
Description: ALE Filing
Official Deadline: March 2, 2026
─────────────────────────
Agent Email: mrgaborkis@gmail.com
Contact Email: hr@globex.com
Contact Phone: tel:+18136019216
Employee Count: 125
─────────────────────────
```

All events are tagged with the **Orange Category** in Outlook so they are easy to spot on your calendar at a glance.

---

## The Completion Summary

When the sync finishes a message box appears showing exactly what happened:

```
Sync Complete!

  Removed:  4 old calendar entries
  Created:  4 new calendar entries

  Breakdown:
  Deadlines :  1
  -90d Alerts:  1
  -30d Alerts:  1
  -7d Alerts:  1
```

- **Removed** = old entries from the previous sync, now deleted and replaced
- **Created** = total new entries added this run
- **Breakdown** = exactly which types were created so you can verify it matches what you expected

---

## How the Cleanup Works (And Why It Is Safe)

Every calendar event this tool creates has a hidden tag stamped into the event body. When the tool runs again it looks for that hidden tag and **only deletes events it created itself**.

This means:
- ✅ Events created by this tool → removed and replaced with fresh ones
- ✅ Your own manually created calendar events → completely safe, never touched
- ✅ Works correctly whether Outlook is open or closed

---

## Common Questions

**Q: I set a client to "No" but their old reminders are still on the calendar. Will they be removed next sync?**
Yes. The cleanup step removes ALL previously synced entries first, then only creates new ones for currently enabled clients. Disabling a client effectively removes all their reminders on the next run.

**Q: What if a deadline falls on a Saturday or Sunday?**
The tool automatically moves it to the Friday before. Same goes for any date that lands on a holiday in your Holidays sheet.

**Q: Can I add more than 3 events per client?**
Not currently. The tool supports up to 3 event columns per client row.
You will need to edit the script line:
  For evBase = 7 To 17 Step 5
and add 5 more columns per event starting at column "V" so if adding 2 more event with total of 10 columns
the new script line will be
  For evBase = 7 To 27 Step 5

**Q: What if I leave the Alerts column blank?**
The tool will still create the DEADLINE event. It just won't create any advance alerts for it.

**Q: The sync button does nothing when I click it.**
This almost always means macros are not enabled. Go back to **Step 1** in the setup section above and make sure you clicked **Enable Content** on the yellow bar after opening the file.

**Q: I see an error that says "User-defined type not defined" or "Can't create object".**
This means the VBA References are not enabled. Go back to **Step 2** in the setup section and make sure both references are checked.

**Q: The sync says it created 0 entries.**
Check that:
1. Column A says `Yes` for at least one client
2. That client's event Scheduled column (G, L, or Q) also says `Yes`
3. The deadline date is a valid date and has not already passed

---

## Quick Reference Card

| Column | What it does | What to type |
|--------|-------------|-------------|
| A | Master on/off for the whole client row | `Yes` or `No` |
| B | Client name | Any text |
| C | Employee count | A number |
| D | Agent email (receives meeting invite) | Email address |
| E | Client contact email | Email address |
| F | Client contact phone | Any phone format |
| G / L / Q | Individual event on/off switch | `Yes` or `No` |
| H / M / R | Event type (short label) | Any text |
| I / N / S | Event description | Any text |
| J / O / T | Alert schedule | e.g. `90d:Y \| 30d:Y \| 7d:Y` |
| K / P / U | Deadline date | e.g. `3/2/2026` |

---

## Setup Checklist (Save This for New Users)

- [ ] macros enabled in Excel (File → Options → Trust Center → Macro Settings)
- [ ] clicked **Enable Content** on the yellow bar when opening the file
- [ ] VBA Reference: **Microsoft Outlook 16.0 Object Library** checked (Alt+F11 → Tools → References)
- [ ] VBA Reference: **Microsoft Scripting Runtime** checked (Alt+F11 → Tools → References)
- [ ] Outlook is installed and signed in with your email account
- [ ] Client data starts on **row 5** of the sheet

---

*Last updated: February 2026 | Compatible with Microsoft 365 and Office 2016–2021 | Windows only*
