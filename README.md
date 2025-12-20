# MVR Ticket Tracker - Reporting System

## Overview

**MVR Ticket Tracker** is a Google Sheets-based reporting system that pulls Freshdesk tickets created by MVR automation emails and maintains a historical record for partner-level reporting.

### Key Features

✅ **Automatic Ticket Fetching** - Pulls tickets matching MVR email pattern from Freshdesk  
✅ **10-Day Lookback Window** - Fetches tickets from last 10 days on each run  
✅ **Historical Append-Only Record** - Maintains complete history without duplicates  
✅ **Subject Line Parsing** - Extracts Type (SC/RC), Partner, and Turn_ID from subject  
✅ **Complete Enrichment** - Adds agent names, status names, resolution times, SLA status  
✅ **Partner Grouping Reports** - Summary statistics grouped by partner  
✅ **Daily Automation** - Optional daily triggers for hands-free operation  

---

## Email Pattern Recognition

The system identifies MVR tickets by subject line pattern:

```
"One Off Continuous MVR - (Suspension Check|Recheck) - {partner} - {turnId}"
```

**Examples:**
- `One Off Continuous MVR - Suspension Check - ABC Transport - T12345`
- `One Off Continuous MVR - Recheck - XYZ Logistics - T67890`

**Extracted Data:**
- **Type**: SC (Suspension Check) or RC (Recheck)
- **Partner**: Partner name
- **Turn_ID**: Turn identifier

---

## Project Structure

### Core Script Files (6)

| File | Purpose | Lines |
|------|---------|-------|
| `Config.gs` | Configuration constants, sheet headers, credentials setup | ~450 |
| `FreshdeskAPI.gs` | Freshdesk API integration with retry logic | ~350 |
| `Enrichment.gs` | Ticket enrichment (names, times, SLA) | ~400 |
| `DataProcessing.gs` | Deduplication, append logic | ~450 |
| `ReportGeneration.gs` | Partner summary reports | ~350 |
| `Code.gs` | Main menu, orchestration, triggers | ~550 |

### Google Sheets Created

1. **MVR_Ticket_History** - Complete historical record of all MVR tickets
2. **By_Partner_Summary** - Partner-level summary statistics
3. **Agent-Mappings** (optional) - Agent ID to name mapping
4. **Freshdesk-Mappings** (optional) - Group/status/priority name mappings

---

## Setup Instructions

### Step 1: Copy Files to Google Apps Script

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete default `Code.gs` content
4. Create 6 new script files with the provided code:
   - Config.gs
   - FreshdeskAPI.gs
   - Enrichment.gs
   - DataProcessing.gs
   - ReportGeneration.gs
   - Code.gs
5. Save all files (**Ctrl+S** or **Cmd+S**)

### Step 2: Configure Freshdesk API

1. Refresh your Google Sheet (close and reopen)
2. Look for the **📊 MVR Tracker** menu (may take 10-15 seconds to appear)
3. Click **📊 MVR Tracker → ⚙️ Setup → 1️⃣ Configure Freshdesk API**
4. Enter your Freshdesk API key when prompted
5. Enter your Freshdesk domain (e.g., "turnhq" for turnhq.freshdesk.com)

**Finding Your API Key:**
- Log into Freshdesk
- Go to Profile Settings → API Key
- Copy the key (looks like: `8NjVDfv7ENvMCHbsG9id`)

### Step 3: Test Connection

1. Click **📊 MVR Tracker → ⚙️ Setup → 2️⃣ Test API Connection**
2. Authorize the script when prompted (first time only)
3. You should see "✅ Freshdesk API connection successful!"

### Step 4: First Run

1. Click **📊 MVR Tracker → 🔄 Data Operations → 🔄 Full Update (Fetch + Reports)**
2. Wait 2-5 minutes for processing
3. Check the results:
   - **MVR_Ticket_History** sheet: All historical tickets
   - **By_Partner_Summary** sheet: Partner statistics

### Step 5: Setup Automatic Updates (Optional)

1. Click **📊 MVR Tracker → ⚙️ Setup → 3️⃣ Setup Automatic Triggers**
2. Confirm to create daily trigger at 8:00 AM CST
3. System will now update automatically every day

---

## Usage Guide

### Manual Operations

| Menu Option | Description | Duration |
|-------------|-------------|----------|
| **Fetch & Append New Tickets** | Pull new tickets from last 10 days | 2-5 min |
| **Update Existing Tickets** | Refresh status for open/pending tickets | 3-10 min |
| **Generate Reports** | Create partner summary from history | <1 min |
| **Full Update** | Fetch + Reports in one step | 2-5 min |

### Viewing Data

- **MVR_Ticket_History**: Complete historical record (append-only)
- **By_Partner_Summary**: Partner-level statistics and metrics
- **View Statistics**: Quick stats popup from menu

### Automatic Operation

With triggers enabled:
- **Daily at 8:00 AM CST**: Fetches new tickets and generates reports
- **Zero manual intervention** required
- **Error notifications** sent to your email if failures occur

---

## Data Sheets Reference

### MVR_Ticket_History

Complete historical record of all MVR tickets with 24 columns:

| Column | Description | Example |
|--------|-------------|---------|
| Ticket_ID | Freshdesk ticket ID | 12345 |
| Subject | Full ticket subject | "One Off Continuous MVR - SC - ABC - T123" |
| Type | SC or RC | SC |
| Partner | Partner name | ABC Transport |
| Turn_ID | Turn identifier | T12345 |
| Status | Status code (2=Open, 4=Resolved) | 2 |
| Status_Name | Human-readable status | Open |
| Priority | Priority code (3=High) | 3 |
| Priority_Name | Human-readable priority | High |
| Created | Creation timestamp | 2025-11-20 08:30 |
| Updated | Last update timestamp | 2025-11-20 10:15 |
| Resolved | Resolution timestamp (if resolved) | 2025-11-20 15:45 |
| Age_Hours | Hours since creation | 12.5 |
| Resolution_Hours | Hours to resolve (if resolved) | 7.2 |
| Agent_ID | Assigned agent ID | 101 |
| Agent_Name | Agent name | John Smith |
| Group_ID | Assigned group ID | 5 |
| Group_Name | Group name | MVR Team |
| Requester_ID | Customer requester ID | 202 |
| Company_ID | Company ID | 303 |
| Tags | Ticket tags | mvr, urgent |
| Source | Source code (1=Email) | 1 |
| Source_Name | Source name | Email |
| SLA | SLA status (MET/BREACH/PEND/Ongoing) | MET |
| Last_Pulled | Last fetch timestamp | 2025-11-20 16:00 |

### By_Partner_Summary

Partner-level summary statistics with 15 columns:

| Column | Description |
|--------|-------------|
| Partner | Partner name |
| Total_Tickets | Total tickets for partner |
| SC_Count | Suspension Check ticket count |
| RC_Count | Recheck ticket count |
| Open_Count | Currently open tickets |
| Pending_Count | Currently pending tickets |
| Resolved_Count | Resolved tickets |
| Closed_Count | Closed tickets |
| Avg_Resolution_Hours | Average resolution time (hours) |
| Avg_Age_Hours | Average ticket age (hours) |
| SLA_Met | Count of tickets meeting SLA |
| SLA_Breach | Count of tickets breaching SLA |
| Oldest_Created | Oldest ticket creation date |
| Newest_Created | Newest ticket creation date |
| Last_Updated | Report generation timestamp |

---

## Configuration Options

### Time Window

Default: 10 days lookback

To change, edit in `Config.gs`:
```javascript
const TIME_CONFIG = {
  LOOKBACK_DAYS: 10,        // Change to 7, 14, 30, etc.
  LOOKBACK_HOURS: 240,      // Update accordingly (days × 24)
  TIMEZONE: "America/Chicago"
};
```

### SLA Thresholds

Default: 24-hour SLA standard

To change, edit in `Config.gs`:
```javascript
const PERFORMANCE_THRESHOLDS = {
  RESOLUTION_STANDARD: 24,  // Change to 48, 72, etc.
  EXCELLENT: 6,             // Adjust as needed
  GOOD: 12,
  WARNING: 18,
  CRITICAL: 24
};
```

### API Rate Limiting

Default: 500ms delay between requests

To change, edit in `Config.gs`:
```javascript
const API_CONFIG = {
  PER_PAGE: 100,
  RATE_LIMIT_DELAY: 500,    // Change to 300, 1000, etc.
  RETRY_ATTEMPTS: 3
};
```

### Trigger Schedule

Default: Daily at 8:00 AM CST

To change:
1. Delete existing trigger: **📊 MVR Tracker → 🗑️ Maintenance → Delete All Triggers**
2. Edit trigger time in `Code.gs`:
```javascript
ScriptApp.newTrigger('dailyUpdate')
  .timeBased()
  .atHour(8)              // Change to 6, 9, 12, etc.
  .everyDays(1)
  .create();
```
3. Re-run: **📊 MVR Tracker → ⚙️ Setup → Setup Automatic Triggers**

---

## Testing & Validation

### Test Functions Available

| Test | Menu Location | Purpose |
|------|---------------|---------|
| API Connection | Setup → Test API Connection | Verify Freshdesk connectivity |
| MVR Ticket Fetch | Testing & Debug → Test MVR Ticket Fetch | Fetch last 24h (quick test) |
| Subject Parsing | Testing & Debug → Test Subject Parsing | Verify pattern matching |
| Enrichment | Testing & Debug → Test Enrichment | Check data enrichment |

### Validation Checklist

- [ ] API credentials configured correctly
- [ ] API connection test passes
- [ ] MVR tickets found in test fetch
- [ ] Subject line parsing extracts Type/Partner/Turn_ID correctly
- [ ] History sheet created with all tickets
- [ ] Partner summary generated with statistics
- [ ] All columns populated (no empty critical fields)
- [ ] Deduplication working (no duplicate Ticket_IDs)
- [ ] SLA classifications correct (MET/BREACH/PEND/Ongoing)

---

## Troubleshooting

### Issue: "Freshdesk credentials not configured"

**Solution:** Run **📊 MVR Tracker → ⚙️ Setup → Configure Freshdesk API**

### Issue: "API connection failed"

**Possible causes:**
- Incorrect API key
- Incorrect domain (should be just "turnhq", not full URL)
- Network/firewall blocking Freshdesk API

**Solution:** Re-run setup with correct credentials

### Issue: "No MVR tickets found"

**Possible causes:**
- No MVR tickets created in last 10 days
- Subject line pattern doesn't match expected format
- Tickets have different subject pattern

**Solution:** Check actual ticket subjects in Freshdesk, adjust pattern if needed

### Issue: Trigger not running automatically

**Check:**
1. **View Active Triggers**: **📊 MVR Tracker → 🗑️ Maintenance → View Active Triggers**
2. Check Apps Script execution log: **Extensions → Apps Script → Executions**

**Solution:** Delete and recreate trigger if needed

### Issue: Duplicate tickets in history

**This should not happen** due to deduplication logic. If it does:
1. Check execution log for errors
2. Verify Ticket_ID column is first column
3. Re-run data processing

---

## Performance & Limits

### Expected Performance

| Operation | Duration | API Calls |
|-----------|----------|-----------|
| Fetch 100 tickets | ~30s | ~10 |
| Fetch 1000 tickets | ~5 min | ~100 |
| Enrich tickets | <1s per 100 | 0 |
| Generate reports | <1s | 0 |

### Freshdesk API Limits

- **Rate limit**: 1000 requests/hour
- **Per page**: 100 tickets maximum
- **Daily quota**: Typically unlimited for paid plans

### Google Apps Script Limits

- **Execution time**: 6 minutes per execution (standard accounts)
- **Daily execution time**: 90 minutes total (standard accounts)
- **Trigger executions**: 20 per day (standard accounts)

**Note:** With 10-day lookback fetching ~700 tickets (70/day × 10), system well within limits.

---

## Maintenance

### Regular Tasks

**Weekly:**
- Review partner summary for data quality
- Check trigger execution log for errors

**Monthly:**
- Verify ticket counts match expectations
- Archive old data if sheet becomes large (>50,000 rows)

**As Needed:**
- Update agent/group mappings if new agents/groups added
- Adjust lookback window if needed
- Update SLA thresholds based on business requirements

### Data Retention

**Current approach:** Unlimited historical retention (append-only)

**If sheet becomes too large (>100,000 rows):**

1. Archive old data to separate sheet
2. Clear history beyond 90 days
3. Adjust lookback window to prevent re-fetching

---

## Advanced Usage

### Custom Reporting

Add custom sheets/functions in `ReportGeneration.gs`:

```javascript
function generateCustomReport() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  const data = historySheet.getDataRange().getValues();
  // Your custom logic here
}
```

### Filtering by Date Range

Modify `DataProcessing.gs` to filter by specific dates:

```javascript
function fetchTicketsByDateRange(startDate, endDate) {
  // Custom implementation
}
```

### Integration with Other Systems

Export data for external analysis:

```javascript
function exportToCSV() {
  const historySheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  // Export logic
}
```

---

## Support & Contact

### Getting Help

1. Check execution logs: **Extensions → Apps Script → Executions**
2. Review error messages in menu alerts
3. Check email for error notifications (if triggers enabled)

### Documentation

- This README
- Inline code comments in all `.gs` files
- Menu documentation: **📊 MVR Tracker → 📚 Documentation**

---

## Version History

**v1.0.0** (November 2025)
- Initial release
- 10-day lookback window
- Automatic deduplication
- Partner grouping reports
- Daily automation support

---

## License & Usage

This system is provided as-is for internal reporting purposes. Modify and adapt as needed for your use case.

---

**Need modifications or have questions?** Check the code comments in each `.gs` file for detailed implementation notes.
