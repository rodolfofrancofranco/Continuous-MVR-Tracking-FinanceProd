# MVR Ticket Tracker - User Manual

## Table of Contents
- [Overview](#overview)
- [Phase 1: Initial Setup](#phase-1-initial-setup-one-time)
- [Phase 2: Backfill Historical Data](#phase-2-backfill-historical-data-one-time)
- [Phase 3: Generate Reports](#phase-3-generate-reports)
- [Phase 4: Manual Override Flow](#phase-4-manual-override-flow)
- [Phase 5: Automated Triggers](#phase-5-automated-triggers)
- [Phase 6: Data Export](#phase-6-data-export)
- [Quick Reference](#quick-reference-daily-operations)
- [Troubleshooting](#troubleshooting)

---

## Overview

The MVR Ticket Tracker automates the tracking and classification of Motor Vehicle Record (MVR) tickets from Freshdesk. It provides:

- **Dynamic Tag Mapping**: Configure tag-to-outcome rules without code changes
- **Outcome Classification**: Automatic classification with billable flag
- **Manual Overrides**: Edit outcomes with full audit trail
- **Stakeholder Reports**: Executive, Finance, and Ops dashboards
- **Data Export**: CSV exports for analysts and auditors

---

## Phase 1: Initial Setup (One-Time)

### Step 1: Open the Google Sheet
- Navigate to your MVR Tracker spreadsheet
- Wait for the custom menu `📊 MVR Tracker` to appear (takes 2-3 seconds)

### Step 2: Initialize Tag Mappings Sheet
```
📊 MVR Tracker → 🏷️ Tag Mappings → 📋 View Tag Mappings
```
- This creates the `Tag_Outcome_Mappings` sheet with 12 default rules
- Verify the sheet has proper dropdowns and checkboxes
- Review/customize mappings if needed (adjust priorities, add new patterns)

### Step 3: Test Tag Matching
```
📊 MVR Tracker → 🏷️ Tag Mappings → 🧪 Test Tag Match
```

**Test 1:**
- Enter: `suspended, reviewed`
- Expected result: Outcome = "Suspension Confirmed", Billable = true

**Test 2:**
- Enter: `valid`
- Expected result: Outcome = "Clear", Billable = false

---

## Phase 2: Backfill Historical Data (One-Time)

### Step 4: Run Backfill
```
📊 MVR Tracker → 🗑️ Maintenance → Backfill Historical Data
```

This process:
- Processes all existing tickets in batches of 50
- Fetches conversations for each ticket
- Applies new outcome classification using dynamic tag mappings
- Populates all 17 new columns (Vendor Group, DL State, Outcomes, etc.)

**Duration**: ~2-5 minutes depending on ticket count

> 💡 Check `View → Execution Log` for progress

### Step 5: Verify Backfill Results
1. Open `MVR_Ticket_History` sheet
2. Scroll right to verify new columns are populated:

| Column | Field |
|--------|-------|
| 26 | Vendor Group |
| 35 | MVR Outcome |
| 41 | Outcome Override |

3. Check for any rows with `Unknown` outcome (may need new tag mappings)

### Step 6: Apply Override Validation
```
📊 MVR Tracker → 📈 View Statistics → Refresh Override Highlighting
```
- Applies dropdown validation to `Outcome Override` column
- Adds yellow highlighting to any rows with manual overrides

---

## Phase 3: Generate Reports

### Step 7: Generate Stakeholder Reports
```
📊 MVR Tracker → 📊 Stakeholder Reports → Generate All Reports
```

Creates/updates:
| Sheet | Purpose |
|-------|---------|
| `Executive_Dashboard` | KPIs and trends |
| `Finance_Audit_Trail` | Billable tickets with full details |
| `Ops_Performance` | Agent metrics |

### Step 8: Generate Monthly Report
```
📊 MVR Tracker → 📊 Generate Reports
```
- Select month/year
- Verify `MVR Requested` column uses dynamic billable logic

### Step 9: Check Override Audit Log
```
📊 MVR Tracker → 📈 View Statistics → View Audit Log
```
- Opens `Override_Audit_Log` sheet
- Will be empty initially (logs future overrides)

---

## Phase 4: Manual Override Flow

### Step 10: Test Override Workflow

1. Open `MVR_Ticket_History`
2. Find a ticket with `MVR Outcome` = "Unknown" or "Still Processing"
3. In `Outcome Override` column (Column 41), select from dropdown: `Clear`
4. In `Override Reason` column (Column 42), enter: `Manual verification - license valid`
5. Run: `📊 MVR Tracker → ⚡ Full Update`
6. Verify:
   - ✅ Row is highlighted yellow
   - ✅ Override preserved (not overwritten)
   - ✅ Entry appears in `Override_Audit_Log`

---

## Phase 5: Automated Triggers

### Step 11: Verify/Create Triggers
```
📊 MVR Tracker → ⚙️ Settings → Manage Triggers
```

Or manually in Apps Script:
1. Open Script Editor (Extensions → Apps Script)
2. Click ⏰ Triggers (left sidebar)
3. Verify these triggers exist:

| Function | Trigger Type | Frequency |
|----------|--------------|-----------|
| `runDailyUpdate` | Time-driven | Daily (6 AM) |
| `runHourlyFetch` | Time-driven | Every hour |

---

## Phase 6: Data Export

### Step 12: Test Exports

**Raw Data Export:**
```
📊 MVR Tracker → 📤 Data Export → Raw Data (CSV)
```
- Creates timestamped CSV in Google Drive
- Verify all 42+ columns are included

**Audit Log Export:**
```
📊 MVR Tracker → 📤 Data Export → Audit Log
```
- Exports override audit trail for compliance

---

## Quick Reference: Daily Operations

| Task | Menu Path | When |
|------|-----------|------|
| Fetch new tickets | `🔄 Fetch & Append New Tickets` | Automatic (hourly) or manual |
| Full update | `⚡ Full Update` | Automatic (daily) or after bulk changes |
| Add tag mapping | `🏷️ Tag Mappings → ➕ Add` | When new outcome pattern identified |
| Manual override | Edit cell in `Outcome Override` column | Edge cases only |
| Generate reports | `📊 Stakeholder Reports → Generate All` | Weekly/Monthly |
| Export for analysts | `📤 Data Export → Raw Data` | On request |

---

## Tag Mappings Sheet Reference

### Columns

| Column | Description | Values |
|--------|-------------|--------|
| Tag Pattern | Text or regex to match | e.g., `suspended`, `dmv[\s_-]?down` |
| Match Type | How to match | `contains` or `regex` |
| Outcome Type | Classification result | Dropdown: Suspension Confirmed, Clear, etc. |
| Priority | Order of evaluation | 1-99 (lower = higher priority) |
| Is Billable | Affects finance reports | Checkbox ✅/❌ |
| Is Active | Include in matching | Checkbox ✅/❌ |
| Notes | Description | Free text |

### Default Mappings

| Pattern | Match Type | Outcome | Priority | Billable |
|---------|------------|---------|----------|----------|
| suspended | contains | Suspension Confirmed | 1 | ✅ |
| valid | contains | Clear | 10 | ❌ |
| active | contains | Clear | 10 | ❌ |
| clear | contains | Clear | 10 | ❌ |
| dmv[\s_-]?down | regex | DMV Unavailable | 20 | ❌ |
| state[\s_-]?down | regex | DMV Unavailable | 20 | ❌ |
| no[\s_-]?dl | regex | Cannot Process | 30 | ❌ |
| discrepancy | contains | Cannot Process | 30 | ❌ |
| missing | contains | Cannot Process | 31 | ❌ |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Menu not appearing | Refresh page, wait 5 seconds |
| "Unknown" outcomes | Add new tag pattern in `Tag_Outcome_Mappings` |
| Backfill timeout | Run again - it continues from where it stopped |
| Override not saved | Check dropdown is valid value from list |
| Missing conversations | API rate limit - wait 1 minute, retry |
| Tag mapping not working | Check `Is Active` checkbox is TRUE |
| Regex not matching | Test pattern at regex101.com first |

---

## Support

For issues or feature requests, contact your system administrator or review the execution logs:
- **Apps Script**: Extensions → Apps Script → View → Executions
- **Logs**: Extensions → Apps Script → View → Logs

---

*Last updated: December 2024*
