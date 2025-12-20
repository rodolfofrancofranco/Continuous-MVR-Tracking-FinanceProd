/**
 * REPORT GENERATION SYSTEM
 * MVR TICKET TRACKER - Partner Grouping & Summary Reports
 * 
 * Purpose: Generate partner-level summary reports from historical ticket data
 */

// ═══════════════════════════════════════════════════════════════════════════════
// PARTNER SUMMARY GENERATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate partner summary report from history
 * Groups tickets by partner and calculates statistics
 * 
 * @return {Object} Summary with partner statistics
 */
function generatePartnerSummary() {
  Logger.log('\n📊 Generating partner summary report...\n');
  
  try {
    // Get history sheet
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found. Run fetchAndAppendNewTickets() first.');
    }
    
    const data = historySheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log('⚠️ No ticket data in history');
      return { partners: [], total: 0 };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Group by partner
    const partnerMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const partner = data[i][colMap['Partner']] || 'Unknown';
      const type = data[i][colMap['Type']];
      const status = data[i][colMap['Status']];
      const statusName = data[i][colMap['Status_Name']];
      const created = new Date(data[i][colMap['Created']]);
      const resolutionHours = parseFloat(data[i][colMap['Resolution_Hours']]) || null;
      const ageHours = parseFloat(data[i][colMap['Age_Hours']]) || 0;
      const sla = data[i][colMap['SLA']];
      
      // Initialize partner entry if not exists
      if (!partnerMap.has(partner)) {
        partnerMap.set(partner, {
          partner: partner,
          total: 0,
          sc_count: 0,
          rc_count: 0,
          open_count: 0,
          pending_count: 0,
          resolved_count: 0,
          closed_count: 0,
          resolution_times: [],
          ages: [],
          sla_met: 0,
          sla_breach: 0,
          oldest_created: null,
          newest_created: null
        });
      }
      
      const entry = partnerMap.get(partner);
      
      // Update counts
      entry.total++;
      
      if (type === 'SC') entry.sc_count++;
      if (type === 'RC') entry.rc_count++;
      
      if (status === TICKET_STATUS.OPEN) entry.open_count++;
      if (status === TICKET_STATUS.PENDING) entry.pending_count++;
      if (status === TICKET_STATUS.RESOLVED) entry.resolved_count++;
      if (status === TICKET_STATUS.CLOSED) entry.closed_count++;
      
      // Track resolution times
      if (resolutionHours !== null) {
        entry.resolution_times.push(resolutionHours);
      }
      
      // Track ages
      entry.ages.push(ageHours);
      
      // Track SLA
      if (sla === 'MET') entry.sla_met++;
      if (sla === 'BREACH') entry.sla_breach++;
      
      // Track date range
      if (!entry.oldest_created || created < entry.oldest_created) {
        entry.oldest_created = created;
      }
      if (!entry.newest_created || created > entry.newest_created) {
        entry.newest_created = created;
      }
    }
    
    // Calculate averages and convert to array
    const partners = Array.from(partnerMap.values()).map(entry => {
      const avgResolution = entry.resolution_times.length > 0
        ? entry.resolution_times.reduce((a, b) => a + b, 0) / entry.resolution_times.length
        : 0;
      
      const avgAge = entry.ages.length > 0
        ? entry.ages.reduce((a, b) => a + b, 0) / entry.ages.length
        : 0;
      
      return {
        ...entry,
        avg_resolution_hours: avgResolution,
        avg_age_hours: avgAge
      };
    });
    
    // Sort by total tickets (descending)
    partners.sort((a, b) => b.total - a.total);
    
    Logger.log(`✅ Generated summary for ${partners.length} partners`);
    
    return {
      partners: partners,
      total: partners.length,
      generated: new Date()
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating partner summary: ${e.message}`);
    throw e;
  }
}

/**
 * Write partner summary to sheet
 * @param {Object} summary - Partner summary object from generatePartnerSummary()
 */
function writePartnerSummary(summary) {
  Logger.log('\n📝 Writing partner summary to sheet...\n');
  
  try {
    // Get or create summary sheet
    const summarySheet = getOrCreateSheet(SHEET_NAMES.BY_PARTNER_SUMMARY, PARTNER_SUMMARY_HEADERS);
    
    // Clear existing data (keep headers)
    const lastRow = summarySheet.getLastRow();
    if (lastRow > 1) {
      summarySheet.getRange(2, 1, lastRow - 1, PARTNER_SUMMARY_HEADERS.length).clearContent();
    }
    
    if (summary.partners.length === 0) {
      Logger.log('⚠️ No partner data to write');
      return;
    }
    
    // Convert partners to rows
    const rows = summary.partners.map(p => [
      p.partner,
      p.total,
      p.sc_count,
      p.rc_count,
      p.open_count,
      p.pending_count,
      p.resolved_count,
      p.closed_count,
      p.avg_resolution_hours.toFixed(1),
      p.avg_age_hours.toFixed(1),
      p.sla_met,
      p.sla_breach,
      p.oldest_created,
      p.newest_created,
      summary.generated
    ]);
    
    // Write rows
    summarySheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    
    // Format sheet
    formatPartnerSummarySheet(summarySheet);
    
    Logger.log(`✅ Wrote ${rows.length} partner summaries to sheet`);
    
  } catch (e) {
    Logger.log(`❌ Error writing partner summary: ${e.message}`);
    throw e;
  }
}

/**
 * Format partner summary sheet with colors and styles
 * @param {Sheet} sheet - Partner summary sheet
 */
function formatPartnerSummarySheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  // Header row - bold and colored
  const headerRange = sheet.getRange(1, 1, 1, PARTNER_SUMMARY_HEADERS.length);
  headerRange.setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Format data rows with alternating colors
  const dataRange = sheet.getRange(2, 1, lastRow - 1, PARTNER_SUMMARY_HEADERS.length);
  
  for (let i = 2; i <= lastRow; i++) {
    const color = (i % 2 === 0) ? '#f8f9fa' : '#ffffff';
    sheet.getRange(i, 1, 1, PARTNER_SUMMARY_HEADERS.length).setBackground(color);
  }
  
  // Number formatting with labels
  const avgResolutionCol = PARTNER_SUMMARY_HEADERS.indexOf('Avg Resolution Time (Hours)') + 1;
  const avgAgeCol = PARTNER_SUMMARY_HEADERS.indexOf('Avg Ticket Age (Hours)') + 1;
  
  if (avgResolutionCol > 0) {
    sheet.getRange(2, avgResolutionCol, lastRow - 1, 1)
      .setNumberFormat('0.0')
      .setHorizontalAlignment('center');
  }
  if (avgAgeCol > 0) {
    sheet.getRange(2, avgAgeCol, lastRow - 1, 1)
      .setNumberFormat('0.0')
      .setHorizontalAlignment('center');
  }
  
  // Date formatting
  const oldestCol = PARTNER_SUMMARY_HEADERS.indexOf('Oldest Ticket Date') + 1;
  const newestCol = PARTNER_SUMMARY_HEADERS.indexOf('Newest Ticket Date') + 1;
  const updatedCol = PARTNER_SUMMARY_HEADERS.indexOf('Report Generated') + 1;
  
  if (oldestCol > 0) {
    sheet.getRange(2, oldestCol, lastRow - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm')
      .setHorizontalAlignment('center');
  }
  if (newestCol > 0) {
    sheet.getRange(2, newestCol, lastRow - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm')
      .setHorizontalAlignment('center');
  }
  if (updatedCol > 0) {
    sheet.getRange(2, updatedCol, lastRow - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss')
      .setHorizontalAlignment('center');
  }
  
  // Highlight status columns with colors
  const openCol = PARTNER_SUMMARY_HEADERS.indexOf('Open Tickets') + 1;
  const pendingCol = PARTNER_SUMMARY_HEADERS.indexOf('Pending Tickets') + 1;
  const resolvedCol = PARTNER_SUMMARY_HEADERS.indexOf('Resolved Tickets') + 1;
  const slaMetCol = PARTNER_SUMMARY_HEADERS.indexOf('SLA Met Count') + 1;
  const slaBreachCol = PARTNER_SUMMARY_HEADERS.indexOf('SLA Breach Count') + 1;
  
  if (openCol > 0) {
    sheet.getRange(2, openCol, lastRow - 1, 1).setBackground('#fff3cd'); // Yellow
  }
  if (pendingCol > 0) {
    sheet.getRange(2, pendingCol, lastRow - 1, 1).setBackground('#cfe2ff'); // Blue
  }
  if (resolvedCol > 0) {
    sheet.getRange(2, resolvedCol, lastRow - 1, 1).setBackground('#d4edda'); // Green
  }
  if (slaMetCol > 0) {
    sheet.getRange(2, slaMetCol, lastRow - 1, 1).setBackground('#d4edda').setFontWeight('bold'); // Green
  }
  if (slaBreachCol > 0) {
    sheet.getRange(2, slaBreachCol, lastRow - 1, 1).setBackground('#f8d7da').setFontWeight('bold'); // Red
  }
  
  // Auto-resize columns
  for (let i = 1; i <= PARTNER_SUMMARY_HEADERS.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Freeze header row and first column (Partner Name)
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  
  Logger.log('✅ Partner summary sheet formatted');
}

// ═══════════════════════════════════════════════════════════════════════════════
// COMPLETE REPORT GENERATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate all reports (partner summary + monthly report)
 * Main entry point for reporting
 */
function generateAllReports() {
  Logger.log('\n📊 === GENERATING ALL REPORTS ===\n');
  
  const startTime = new Date();
  const results = {};
  
  try {
    // Generate partner summary
    Logger.log('1️⃣ Generating partner summary...');
    const summary = generatePartnerSummary();
    Logger.log('✅ Partner summary generated\n');
    
    // Write to sheet
    Logger.log('2️⃣ Writing partner summary to sheet...');
    writePartnerSummary(summary);
    Logger.log('✅ Partner summary written\n');
    
    // Generate monthly report
    Logger.log('3️⃣ Generating monthly report...');
    const monthlyReport = generateMonthlyRequestReport();
    Logger.log('✅ Monthly report generated\n');
    
    // Write monthly report
    Logger.log('4️⃣ Writing monthly report to sheet...');
    writeMonthlyReport(monthlyReport);
    Logger.log('✅ Monthly report written\n');
    
    // Generate Assumptions Log
    Logger.log('5️⃣ Generating Assumptions Log...');
    results.assumptions = generateAssumptionsLog();
    Logger.log('✅ Assumptions Log generated\n');
    
    // Generate OPS Lifecycle Report
    Logger.log('6️⃣ Generating OPS Lifecycle Report...');
    results.ops = generateOpsLifecycleReport();
    Logger.log('✅ OPS Lifecycle Report generated\n');
    
    // Generate Finance Reconciliation Report
    Logger.log('7️⃣ Generating Finance Reconciliation Report...');
    results.finance = generateFinanceReconciliationReport();
    Logger.log('✅ Finance Reconciliation Report generated\n');
    
    // Generate CEO Effectiveness Report
    Logger.log('8️⃣ Generating CEO Effectiveness Report...');
    results.ceo = generateCEOEffectivenessReport();
    Logger.log('✅ CEO Effectiveness Report generated\n');
    
    // Generate Annual Projection
    Logger.log('9️⃣ Generating Annual Projection...');
    results.annual = generateAnnualProjection();
    Logger.log('✅ Annual Projection generated\n');
    
    // Generate Operative Plan
    Logger.log('🔟 Generating Operative Plan...');
    results.operative = generateOperativePlan();
    Logger.log('✅ Operative Plan generated\n');
    
    // Generate Growth Projection
    Logger.log('1️⃣1️⃣ Generating Growth Projection...');
    results.growth = generateGrowthProjection();
    Logger.log('✅ Growth Projection generated\n');
    
    const duration = (new Date() - startTime) / 1000;
    
    Logger.log('\n📊 === REPORT GENERATION COMPLETE ===');
    Logger.log(`   Partners: ${summary.total}`);
    Logger.log(`   Months: ${monthlyReport.total}`);
    Logger.log(`   New Reports: Assumptions, OPS, Finance, CEO, Annual, Operative, Growth`);
    Logger.log(`   Duration: ${duration.toFixed(1)}s`);
    Logger.log('\n✅ All reports generated!\n');
    
    return {
      success: true,
      partners: summary.total,
      months: monthlyReport.total,
      results: results,
      duration: duration
    };
    
  } catch (e) {
    Logger.log(`\n❌ Error generating reports: ${e.message}`);
    Logger.log(e.stack);
    
    return {
      success: false,
      error: e.message,
      results: results,
      duration: (new Date() - startTime) / 1000
    };
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// QUICK STATS & SUMMARIES
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Get quick statistics for display
 * @return {Object} Quick stats object
 */
function getQuickStats() {
  try {
    const historyStats = getHistoryStats();
    
    if (historyStats.error) {
      return historyStats;
    }
    
    const summarySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BY_PARTNER_SUMMARY);
    const partnerCount = summarySheet ? summarySheet.getLastRow() - 1 : 0;
    
    return {
      total_tickets: historyStats.total,
      by_type: historyStats.by_type,
      by_status: historyStats.by_status,
      partners: partnerCount,
      date_range: {
        oldest: historyStats.oldest_created,
        newest: historyStats.newest_created
      }
    };
    
  } catch (e) {
    Logger.log(`❌ Error getting quick stats: ${e.message}`);
    return { error: e.message };
  }
}

/**
 * Display quick stats in UI
 */
function showQuickStats() {
  const stats = getQuickStats();
  
  if (stats.error) {
    SpreadsheetApp.getUi().alert('Error', stats.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const message = `
📊 MVR TICKET TRACKER STATISTICS

Total Tickets: ${stats.total_tickets}

By Type:
  • SC (Suspension Check): ${stats.by_type.SC || 0}
  • RC (Recheck): ${stats.by_type.RC || 0}

By Status:
  • Open: ${stats.by_status?.Open || 0}
  • Pending: ${stats.by_status?.Pending || 0}
  • Resolved: ${stats.by_status?.Resolved || 0}
  • Closed: ${stats.by_status?.Closed || 0}

By Vendor Group:
  • CERTN: ${stats.by_vendor_group?.CERTN || 0}
  • PENNDOT: ${stats.by_vendor_group?.PENNDOT || 0}

By Outcome:
  • Suspension Confirmed: ${stats.by_outcome?.['Suspension Confirmed'] || 0}
  • Clear: ${stats.by_outcome?.Clear || 0}
  • Still Processing: ${stats.by_outcome?.['Still Processing'] || 0}

Manual Overrides: ${stats.overrides_count || 0}
Partners: ${stats.partners}

Date Range:
  ${stats.date_range?.oldest?.toLocaleDateString()} to ${stats.date_range?.newest?.toLocaleDateString()}
  `;
  
  SpreadsheetApp.getUi().alert('MVR Ticket Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// EXECUTIVE DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Executive Dashboard with high-level KPIs
 * @return {Object} Dashboard data
 */
function generateExecutiveDashboard() {
  Logger.log('\n📊 Generating Executive Dashboard...\n');
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found');
    }
    
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: 'No ticket data' };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Calculate KPIs
    const kpis = {
      total_volume: data.length - 1,
      by_status: {},
      by_vendor_group: {},
      by_outcome: {},
      sla_met: 0,
      sla_breach: 0,
      avg_resolution_hours: 0,
      overrides_count: 0,
      overrides_rate: 0,
      billable_count: 0,
      billable_rate: 0,
      // Monthly trend
      monthly_volumes: {},
      generated_at: new Date()
    };
    
    let totalResolutionHours = 0;
    let resolvedCount = 0;
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][colMap['Status']];
      const vendorGroup = data[i][colMap['Vendor Group']];
      const outcome = data[i][colMap['MVR Outcome']];
      const override = data[i][colMap['Outcome Override']];
      const sla = data[i][colMap['SLA Status']];
      const resolutionHours = parseFloat(data[i][colMap['Resolution Time (Hours)']]) || 0;
      const created = new Date(data[i][colMap['Date Created']]);
      
      // Status counts
      if (status) kpis.by_status[status] = (kpis.by_status[status] || 0) + 1;
      
      // Vendor group counts
      if (vendorGroup) kpis.by_vendor_group[vendorGroup] = (kpis.by_vendor_group[vendorGroup] || 0) + 1;
      
      // Outcome counts
      const effectiveOutcome = override || outcome;
      if (effectiveOutcome) kpis.by_outcome[effectiveOutcome] = (kpis.by_outcome[effectiveOutcome] || 0) + 1;
      
      // Override count
      if (override) kpis.overrides_count++;
      
      // SLA counts
      if (sla === 'MET') kpis.sla_met++;
      if (sla === 'BREACH') kpis.sla_breach++;
      
      // Resolution time
      if (resolutionHours > 0) {
        totalResolutionHours += resolutionHours;
        resolvedCount++;
      }
      
      // Billable (suspension confirmed)
      if (effectiveOutcome === OUTCOME_TYPES.SUSPENSION_CONFIRMED) {
        kpis.billable_count++;
      }
      
      // Monthly volume
      if (created && !isNaN(created.getTime())) {
        const monthKey = `${created.getFullYear()}-${String(created.getMonth() + 1).padStart(2, '0')}`;
        kpis.monthly_volumes[monthKey] = (kpis.monthly_volumes[monthKey] || 0) + 1;
      }
    }
    
    // Calculate rates
    kpis.avg_resolution_hours = resolvedCount > 0 ? totalResolutionHours / resolvedCount : 0;
    kpis.overrides_rate = kpis.total_volume > 0 ? (kpis.overrides_count / kpis.total_volume * 100) : 0;
    kpis.billable_rate = kpis.total_volume > 0 ? (kpis.billable_count / kpis.total_volume * 100) : 0;
    kpis.sla_rate = (kpis.sla_met + kpis.sla_breach) > 0 
      ? (kpis.sla_met / (kpis.sla_met + kpis.sla_breach) * 100) 
      : 0;
    
    // Write to dashboard sheet
    writeExecutiveDashboard(kpis);
    
    Logger.log('✅ Executive Dashboard generated');
    return kpis;
    
  } catch (e) {
    Logger.log(`❌ Error generating dashboard: ${e.message}`);
    throw e;
  }
}

/**
 * Write Executive Dashboard to sheet
 */
function writeExecutiveDashboard(kpis) {
  const dashHeaders = [
    "Metric", "Value", "Details", "Updated"
  ];
  
  const sheet = getOrCreateSheet(SHEET_NAMES.EXECUTIVE_DASHBOARD, dashHeaders);
  
  // Clear existing data
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, dashHeaders.length).clearContent();
  }
  
  // Build metrics rows
  const rows = [
    ["Total Volume", kpis.total_volume, "", kpis.generated_at],
    ["", "", "", ""],
    ["=== STATUS BREAKDOWN ===", "", "", ""],
    ["Open Tickets", kpis.by_status?.Open || 0, "", ""],
    ["Pending Tickets", kpis.by_status?.Pending || 0, "", ""],
    ["Resolved Tickets", kpis.by_status?.Resolved || 0, "", ""],
    ["Closed Tickets", kpis.by_status?.Closed || 0, "", ""],
    ["", "", "", ""],
    ["=== SLA PERFORMANCE ===", "", "", ""],
    ["SLA Met", kpis.sla_met, `${kpis.sla_rate.toFixed(1)}%`, ""],
    ["SLA Breach", kpis.sla_breach, "", ""],
    ["Avg Resolution (Hours)", kpis.avg_resolution_hours.toFixed(1), "", ""],
    ["", "", "", ""],
    ["=== VENDOR GROUPS ===", "", "", ""],
    ["CERTN", kpis.by_vendor_group?.CERTN || 0, "", ""],
    ["PENNDOT", kpis.by_vendor_group?.PENNDOT || 0, "", ""],
    ["LEGACY", kpis.by_vendor_group?.LEGACY || 0, "", ""],
    ["", "", "", ""],
    ["=== OUTCOMES ===", "", "", ""],
    ["Suspension Confirmed", kpis.by_outcome?.[OUTCOME_TYPES.SUSPENSION_CONFIRMED] || 0, "Billable", ""],
    ["Clear", kpis.by_outcome?.[OUTCOME_TYPES.CLEAR] || 0, "", ""],
    ["DMV Unavailable", kpis.by_outcome?.[OUTCOME_TYPES.DMV_UNAVAILABLE] || 0, "", ""],
    ["Cannot Process", kpis.by_outcome?.[OUTCOME_TYPES.CANNOT_PROCESS] || 0, "", ""],
    ["Still Processing", kpis.by_outcome?.[OUTCOME_TYPES.STILL_PROCESSING] || 0, "", ""],
    ["", "", "", ""],
    ["=== BILLING ===", "", "", ""],
    ["Billable Count", kpis.billable_count, `${kpis.billable_rate.toFixed(1)}%`, ""],
    ["Manual Overrides", kpis.overrides_count, `${kpis.overrides_rate.toFixed(1)}%`, ""]
  ];
  
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  
  // Format
  formatDashboardSheet(sheet);
}

/**
 * Format dashboard sheet
 */
function formatDashboardSheet(sheet) {
  // Header formatting
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
  
  // Section headers
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().startsWith('===')) {
      sheet.getRange(i + 1, 1, 1, 4).setBackground('#e8f0fe').setFontWeight('bold');
    }
  }
  
  // Auto-resize
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.autoResizeColumn(3);
}

// ═══════════════════════════════════════════════════════════════════════════════
// FINANCE AUDIT TRAIL
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Finance Audit Trail - billable tickets with full details
 */
function generateFinanceAuditTrail() {
  Logger.log('\n💰 Generating Finance Audit Trail...\n');
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found');
    }
    
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: 'No ticket data' };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Audit trail headers
    const auditHeaders = [
      "Ticket ID",
      "Turn ID",
      "Partner",
      "DL State",
      "Vendor Group",
      "Date Created",
      "Date Resolved",
      "MVR Outcome",
      "Is Override",
      "Override Reason",
      "Status",
      "SLA Status",
      "Freshdesk URL",
      "Export Date"
    ];
    
    const auditSheet = getOrCreateSheet(SHEET_NAMES.FINANCE_AUDIT_TRAIL, auditHeaders);
    
    // Clear existing data
    const lastRow = auditSheet.getLastRow();
    if (lastRow > 1) {
      auditSheet.getRange(2, 1, lastRow - 1, auditHeaders.length).clearContent();
    }
    
    // Get Freshdesk domain for URL construction
    let freshdeskDomain = "";
    try {
      const creds = getFreshdeskCredentials();
      freshdeskDomain = creds.domain;
    } catch (e) {
      freshdeskDomain = "unknown";
    }
    
    // Build audit rows
    const auditRows = [];
    const exportDate = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const ticketId = data[i][colMap['Ticket ID']];
      const turnId = data[i][colMap['Turn ID']];
      const partner = data[i][colMap['Partner Name']];
      const dlState = data[i][colMap['DL State']];
      const vendorGroup = data[i][colMap['Vendor Group']];
      const created = data[i][colMap['Date Created']];
      const resolved = data[i][colMap['Date Resolved']];
      const outcome = data[i][colMap['MVR Outcome']];
      const override = data[i][colMap['Outcome Override']];
      const overrideReason = data[i][colMap['Override Reason']];
      const status = data[i][colMap['Status']];
      const sla = data[i][colMap['SLA Status']];
      
      const effectiveOutcome = override || outcome;
      const isOverride = override ? "Yes" : "No";
      
      const freshdeskUrl = `https://${freshdeskDomain}.freshdesk.com/a/tickets/${ticketId}`;
      
      auditRows.push([
        ticketId,
        turnId,
        partner,
        dlState,
        vendorGroup,
        created,
        resolved,
        effectiveOutcome,
        isOverride,
        overrideReason || "",
        status,
        sla,
        freshdeskUrl,
        exportDate
      ]);
    }
    
    if (auditRows.length > 0) {
      auditSheet.getRange(2, 1, auditRows.length, auditHeaders.length).setValues(auditRows);
    }
    
    // Format
    formatAuditSheet(auditSheet, auditHeaders);
    
    Logger.log(`✅ Generated Finance Audit Trail with ${auditRows.length} records`);
    
    return {
      success: true,
      total: auditRows.length,
      generated: exportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating audit trail: ${e.message}`);
    throw e;
  }
}

/**
 * Format audit trail sheet
 */
function formatAuditSheet(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  // Header formatting
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
  
  // Highlight override rows
  const colMap = createColumnMapper(headers);
  const overrideCol = colMap['Is Override'] + 1;
  
  for (let i = 2; i <= lastRow; i++) {
    const isOverride = sheet.getRange(i, overrideCol).getValue();
    if (isOverride === "Yes") {
      sheet.getRange(i, 1, 1, headers.length).setBackground('#FFEB3B');
    }
  }
  
  // Auto-resize
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Freeze header and first columns
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// OPS PERFORMANCE REPORT
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Ops Performance report - by-agent metrics
 */
function generateOpsPerformance() {
  Logger.log('\n👥 Generating Ops Performance Report...\n');
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      throw new Error('History sheet not found');
    }
    
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: 'No ticket data' };
    }
    
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Aggregate by agent
    const agentMap = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const agent = data[i][colMap['Assigned Agent']] || "Unassigned";
      const resolutionHours = parseFloat(data[i][colMap['Resolution Time (Hours)']]) || 0;
      const sla = data[i][colMap['SLA Status']];
      const outcome = data[i][colMap['MVR Outcome']];
      const override = data[i][colMap['Outcome Override']];
      const status = data[i][colMap['Status']];
      
      if (!agentMap.has(agent)) {
        agentMap.set(agent, {
          agent: agent,
          total_tickets: 0,
          resolved_tickets: 0,
          resolution_times: [],
          sla_met: 0,
          sla_breach: 0,
          overrides: 0,
          by_outcome: {}
        });
      }
      
      const entry = agentMap.get(agent);
      entry.total_tickets++;
      
      if (status === 'Resolved' || status === 'Closed') {
        entry.resolved_tickets++;
        if (resolutionHours > 0) {
          entry.resolution_times.push(resolutionHours);
        }
      }
      
      if (sla === 'MET') entry.sla_met++;
      if (sla === 'BREACH') entry.sla_breach++;
      
      if (override) entry.overrides++;
      
      const effectiveOutcome = override || outcome;
      if (effectiveOutcome) {
        entry.by_outcome[effectiveOutcome] = (entry.by_outcome[effectiveOutcome] || 0) + 1;
      }
    }
    
    // Ops Performance headers
    const opsHeaders = [
      "Agent",
      "Total Tickets",
      "Resolved",
      "Avg Resolution (Hours)",
      "SLA Met",
      "SLA Breach",
      "SLA Rate %",
      "Overrides",
      "Override Rate %",
      "Report Date"
    ];
    
    const opsSheet = getOrCreateSheet(SHEET_NAMES.OPS_PERFORMANCE, opsHeaders);
    
    // Clear existing data
    const lastRow = opsSheet.getLastRow();
    if (lastRow > 1) {
      opsSheet.getRange(2, 1, lastRow - 1, opsHeaders.length).clearContent();
    }
    
    // Build rows
    const reportDate = new Date();
    const opsRows = [];
    
    for (const [agent, entry] of agentMap) {
      const avgResolution = entry.resolution_times.length > 0
        ? entry.resolution_times.reduce((a, b) => a + b, 0) / entry.resolution_times.length
        : 0;
      
      const slaRate = (entry.sla_met + entry.sla_breach) > 0
        ? (entry.sla_met / (entry.sla_met + entry.sla_breach) * 100)
        : 0;
      
      const overrideRate = entry.total_tickets > 0
        ? (entry.overrides / entry.total_tickets * 100)
        : 0;
      
      opsRows.push([
        agent,
        entry.total_tickets,
        entry.resolved_tickets,
        avgResolution.toFixed(1),
        entry.sla_met,
        entry.sla_breach,
        slaRate.toFixed(1) + "%",
        entry.overrides,
        overrideRate.toFixed(1) + "%",
        reportDate
      ]);
    }
    
    // Sort by total tickets descending
    opsRows.sort((a, b) => b[1] - a[1]);
    
    if (opsRows.length > 0) {
      opsSheet.getRange(2, 1, opsRows.length, opsHeaders.length).setValues(opsRows);
    }
    
    // Format
    formatOpsSheet(opsSheet, opsHeaders);
    
    Logger.log(`✅ Generated Ops Performance Report for ${opsRows.length} agents`);
    
    return {
      success: true,
      agents: opsRows.length,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating ops performance: ${e.message}`);
    throw e;
  }
}

/**
 * Format ops performance sheet
 */
function formatOpsSheet(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  
  // Header formatting
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');
  
  // Highlight SLA breach column if high
  const colMap = createColumnMapper(headers);
  const slaBreachCol = colMap['SLA Breach'] + 1;
  
  for (let i = 2; i <= lastRow; i++) {
    const breachCount = sheet.getRange(i, slaBreachCol).getValue();
    if (breachCount > 5) {
      sheet.getRange(i, slaBreachCol).setBackground('#f8d7da');
    }
  }
  
  // Auto-resize
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  sheet.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ASSUMPTIONS LOG - Central Reference for All Reports
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Assumptions Log sheet
 * Central reference document for all assumptions used in projections and reports
 * 
 * @return {Object} Summary of generated log
 */
function generateAssumptionsLog() {
  Logger.log('\n📋 === GENERATING ASSUMPTIONS LOG ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.ASSUMPTIONS_LOG;
  const reportDate = new Date().toISOString().split('T')[0];
  
  try {
    // Get or create sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build assumptions data
    const rows = [];
    
    // Header section
    rows.push(['MVR REPORTING SUITE - ASSUMPTIONS LOG', '', '', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', '', '', '', '']);
    rows.push(['Purpose:', 'Document all assumptions used in projections and reports. Update values here to recalculate.', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '']);
    
    // Column headers
    rows.push(['Category', 'Variable', 'Value', 'Low', 'Mid', 'High', 'Description', 'Source/Rationale', 'Last Updated']);
    
    // ─── TIER CONFIGURATION ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['TIER CONFIGURATION', '', '', '', '', '', '', '', '']);
    
    Object.entries(TIER_CONFIG).forEach(([tier, config]) => {
      rows.push([
        '',
        `Tier ${tier} Cadence`,
        `${config.cadence_days} days`,
        '', '', '',
        config.description,
        'Business rule - state risk classification',
        reportDate
      ]);
      rows.push([
        '',
        `Tier ${tier} Checks/Year`,
        config.checks_per_year,
        '', '', '',
        `365 / ${config.cadence_days} = ${config.checks_per_year} checks per enrolled driver per year`,
        'Calculated from cadence',
        reportDate
      ]);
    });
    
    // ─── CAPACITY CONFIGURATION ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['CAPACITY CONFIGURATION', '', '', '', '', '', '', '', '']);
    
    rows.push(['', 'Max Checks Per Day', CAPACITY_CONFIG.MAX_CHECKS_PER_DAY, '', '', '',
      'Maximum tickets an agent can process if doing 100% MVR work',
      'Operational benchmark from team leads', reportDate]);
    
    rows.push(['', 'MVR Time Allocation %', `${CAPACITY_CONFIG.MVR_TIME_ALLOCATION_PCT}%`, '', '', '',
      'Percentage of agent time dedicated to MVR vs other ticket types',
      'Estimate - agents handle multiple ticket types', reportDate]);
    
    rows.push(['', 'Effective Checks Per Day', CAPACITY_CONFIG.EFFECTIVE_CHECKS_PER_DAY, '', '', '',
      `${CAPACITY_CONFIG.MAX_CHECKS_PER_DAY} × ${CAPACITY_CONFIG.MVR_TIME_ALLOCATION_PCT}% = ${CAPACITY_CONFIG.EFFECTIVE_CHECKS_PER_DAY} effective MVR checks/day/agent`,
      'Calculated', reportDate]);
    
    rows.push(['', 'Working Days Per Month', CAPACITY_CONFIG.WORKING_DAYS_PER_MONTH, '', '', '',
      'Average working days per month (excludes weekends/holidays)',
      'Standard assumption', reportDate]);
    
    rows.push(['', 'Working Days Per Year', CAPACITY_CONFIG.WORKING_DAYS_PER_YEAR, '', '', '',
      '21 days × 12 months = 252 working days',
      'Calculated', reportDate]);
    
    // ─── UNCERTAINTY RANGES ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['UNCERTAINTY RANGES', '', '', '', '', '', '', '', '']);
    
    rows.push(['', 'Enrolled Profile %',
      `${ASSUMPTIONS.ENROLLED_PCT.mid}%`,
      `${ASSUMPTIONS.ENROLLED_PCT.low}%`,
      `${ASSUMPTIONS.ENROLLED_PCT.mid}%`,
      `${ASSUMPTIONS.ENROLLED_PCT.high}%`,
      ASSUMPTIONS.ENROLLED_PCT.description,
      ASSUMPTIONS.ENROLLED_PCT.source, reportDate]);
    
    rows.push(['', 'Active Profile %',
      `${ASSUMPTIONS.ACTIVE_PROFILE_PCT.mid}%`,
      `${ASSUMPTIONS.ACTIVE_PROFILE_PCT.low}%`,
      `${ASSUMPTIONS.ACTIVE_PROFILE_PCT.mid}%`,
      `${ASSUMPTIONS.ACTIVE_PROFILE_PCT.high}%`,
      ASSUMPTIONS.ACTIVE_PROFILE_PCT.description,
      ASSUMPTIONS.ACTIVE_PROFILE_PCT.source, reportDate]);
    
    rows.push(['', 'FREE State Driver %',
      `${ASSUMPTIONS.FREE_STATE_PCT.mid}%`,
      `${ASSUMPTIONS.FREE_STATE_PCT.low}%`,
      `${ASSUMPTIONS.FREE_STATE_PCT.mid}%`,
      `${ASSUMPTIONS.FREE_STATE_PCT.high}%`,
      ASSUMPTIONS.FREE_STATE_PCT.description,
      ASSUMPTIONS.FREE_STATE_PCT.source, reportDate]);
    
    // ─── PROBLEM THRESHOLDS ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['PROBLEM THRESHOLDS', '', '', '', '', '', '', '', '']);
    
    rows.push(['', 'Open Hours (Warning)', `${ASSUMPTIONS.PROBLEM_THRESHOLDS.open_hours_warning} hrs`, '', '', '',
      'Flag tickets open longer than this threshold',
      'OPS operational target', reportDate]);
    
    rows.push(['', 'Open Hours (Critical)', `${ASSUMPTIONS.PROBLEM_THRESHOLDS.open_hours_critical} hrs`, '', '', '',
      'Critical alert for tickets open longer than this',
      'OPS escalation threshold', reportDate]);
    
    rows.push(['', 'Pending Hours (Warning)', `${ASSUMPTIONS.PROBLEM_THRESHOLDS.pending_hours_warning} hrs`, '', '', '',
      'Flag tickets pending longer than this threshold',
      'OPS operational target', reportDate]);
    
    rows.push(['', 'Pending Hours (Critical)', `${ASSUMPTIONS.PROBLEM_THRESHOLDS.pending_hours_critical} hrs`, '', '', '',
      'Critical alert for tickets pending longer than this',
      'OPS escalation threshold', reportDate]);
    
    // ─── GROWTH SCENARIOS ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['GROWTH SCENARIOS', '', '', '', '', '', '', '', '']);
    
    rows.push(['', 'Conservative Growth', `${ASSUMPTIONS.GROWTH_SCENARIOS.conservative * 100}%`, '', '', '',
      'Annual enrollment growth rate - conservative estimate',
      'Historical growth patterns', reportDate]);
    
    rows.push(['', 'Moderate Growth', `${ASSUMPTIONS.GROWTH_SCENARIOS.moderate * 100}%`, '', '', '',
      'Annual enrollment growth rate - moderate estimate',
      'Target growth rate', reportDate]);
    
    rows.push(['', 'Aggressive Growth', `${ASSUMPTIONS.GROWTH_SCENARIOS.aggressive * 100}%`, '', '', '',
      'Annual enrollment growth rate - aggressive estimate',
      'Stretch target', reportDate]);
    
    rows.push(['', 'Projection Horizon', `${ASSUMPTIONS.PROJECTION_YEARS} years`, '', '', '',
      'Number of years to project forward',
      'Standard planning horizon', reportDate]);
    
    // ─── STATE CONFIGURATION SUMMARY ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['STATE CONFIGURATION SUMMARY', '', '', '', '', '', '', '', '']);
    
    // Count states by tier and free/paid
    const stateCounts = { tier1: 0, tier2: 0, tier3: 0, free: 0, paid: 0 };
    Object.values(STATE_CONFIG).forEach(config => {
      if (config.tier === 1) stateCounts.tier1++;
      else if (config.tier === 2) stateCounts.tier2++;
      else stateCounts.tier3++;
      if (config.is_free) stateCounts.free++;
      else stateCounts.paid++;
    });
    
    rows.push(['', 'Tier 1 States (60-day)', stateCounts.tier1, '', '', '',
      'DE, MT, MS - High risk states requiring frequent checks',
      'State risk classification', reportDate]);
    
    rows.push(['', 'Tier 2 States (90-day)', stateCounts.tier2, '', '', '',
      'GA, LA, TN, MI - Medium risk states',
      'State risk classification', reportDate]);
    
    rows.push(['', 'Tier 3 States (180-day)', stateCounts.tier3, '', '', '',
      'All remaining states - Standard cadence',
      'State risk classification', reportDate]);
    
    rows.push(['', 'FREE States (DMV Check)', stateCounts.free, '', '', '',
      'States where DMV site check is free (no vendor cost)',
      'Vendor contract terms', reportDate]);
    
    rows.push(['', 'PAID States (Vendor)', stateCounts.paid, '', '', '',
      'States requiring paid vendor check (INFORM, PENNDOT, CERTN)',
      'Vendor contract terms', reportDate]);
    
    // ─── PROBLEM TAGS ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['PROBLEM TAGS (For OPS Reporting)', '', '', '', '', '', '', '', '']);
    
    PROBLEM_TAGS.forEach(tag => {
      rows.push(['', tag, '', '', '', '',
        'Tag indicating resolution problem - flagged in OPS lifecycle report',
        'Freshdesk tag taxonomy', reportDate]);
    });
    
    // ─── DATA SOURCES ───
    rows.push(['', '', '', '', '', '', '', '', '']);
    rows.push(['DATA SOURCES', '', '', '', '', '', '', '', '']);
    
    rows.push(['', 'Primary Source', 'Freshdesk API', '', '', '',
      'All ticket data pulled via Freshdesk API v2',
      'API endpoint: /api/v2/tickets', reportDate]);
    
    rows.push(['', 'Refresh Cadence', 'Hourly + On-Demand', '', '', '',
      'Automated hourly sync plus manual full refresh option',
      'Trigger configuration', reportDate]);
    
    rows.push(['', 'History Retention', 'Append-only', '', '', '',
      'Historical data preserved in MVR_Ticket_History sheet',
      'Data architecture decision', reportDate]);
    
    // Write all rows
    sheet.getRange(1, 1, rows.length, 9).setValues(rows);
    
    // Format the sheet
    formatAssumptionsLog(sheet, rows.length);
    
    Logger.log(`✅ Assumptions Log generated with ${rows.length} rows`);
    
    return {
      success: true,
      rows: rows.length,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating assumptions log: ${e.message}`);
    throw e;
  }
}

/**
 * Format the assumptions log sheet for readability
 */
function formatAssumptionsLog(sheet, totalRows) {
  // Title formatting
  sheet.getRange(1, 1, 1, 9).merge()
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14);
  
  // Subtitle rows
  sheet.getRange(2, 1, 2, 9).setFontStyle('italic').setFontColor('#666666');
  
  // Column headers
  sheet.getRange(5, 1, 1, 9)
    .setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold');
  
  // Category headers (rows with text only in column A)
  for (let i = 7; i <= totalRows; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (cellValue && cellValue !== '' && !sheet.getRange(i, 2).getValue()) {
      sheet.getRange(i, 1, 1, 9)
        .setBackground('#e8f0fe').setFontWeight('bold');
    }
  }
  
  // Auto-resize columns
  for (let i = 1; i <= 9; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Set column widths for better readability
  sheet.setColumnWidth(7, 350);  // Description
  sheet.setColumnWidth(8, 250);  // Source/Rationale
  
  // Freeze header rows
  sheet.setFrozenRows(5);
}

// ═══════════════════════════════════════════════════════════════════════════════
// OPS LIFECYCLE REPORT - Ticket Flow & Resolution Problems
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate OPS Lifecycle Report
 * Focuses on ticket lifecycle, tag patterns, and resolution problems
 * 
 * @return {Object} Summary of generated report
 */
function generateOpsLifecycleReport() {
  Logger.log('\n📊 === GENERATING OPS LIFECYCLE REPORT ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.OPS_PERFORMANCE;
  const reportDate = new Date().toISOString().split('T')[0];
  
  try {
    // Get raw tickets data
    const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      throw new Error('No raw ticket data available. Run initial pull first.');
    }
    
    const data = rawSheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Initialize metrics
    const statusCounts = { Open: 0, Pending: 0, Resolved: 0, Closed: 0, Other: 0 };
    const tagFrequency = {};
    const problemTagCounts = {};
    const statusTagMatrix = {};  // status -> tag -> count
    const ageByStatus = { Open: [], Pending: [] };
    const resolutionTimes = [];
    const blockedTickets = [];  // Tickets with problem tags still open
    const tierMetrics = {};
    
    // Initialize problem tag counts
    PROBLEM_TAGS.forEach(tag => problemTagCounts[tag] = 0);
    
    // Process each ticket
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ticketId = row[colMap['Ticket ID']];
      const status = row[colMap['Status']] || 'Other';
      const tagsRaw = row[colMap['Tags']] || '';
      const createdAt = row[colMap['Created At']];
      const resolvedAt = row[colMap['Resolved At']];
      const tier = row[colMap['Tier']] || 'Unknown';
      
      // Status counts
      if (statusCounts.hasOwnProperty(status)) {
        statusCounts[status]++;
      } else {
        statusCounts.Other++;
      }
      
      // Parse tags
      const tags = tagsRaw.split(',').map(t => t.trim()).filter(t => t);
      
      // Tag frequency
      tags.forEach(tag => {
        tagFrequency[tag] = (tagFrequency[tag] || 0) + 1;
        
        // Check if problem tag
        if (PROBLEM_TAGS.includes(tag)) {
          problemTagCounts[tag]++;
        }
        
        // Status × Tag matrix
        if (!statusTagMatrix[status]) statusTagMatrix[status] = {};
        statusTagMatrix[status][tag] = (statusTagMatrix[status][tag] || 0) + 1;
      });
      
      // Age calculations for open/pending
      if (createdAt && (status === 'Open' || status === 'Pending')) {
        const ageHours = (new Date() - new Date(createdAt)) / (1000 * 60 * 60);
        ageByStatus[status].push({
          ticketId,
          ageHours,
          tags: tagsRaw,
          tier
        });
        
        // Check for blocked tickets (problem tags + still open)
        const hasProblemTag = tags.some(t => PROBLEM_TAGS.includes(t));
        if (hasProblemTag && ageHours > ASSUMPTIONS.PROBLEM_THRESHOLDS.open_hours_warning) {
          blockedTickets.push({
            ticketId,
            status,
            ageHours: Math.round(ageHours),
            problemTags: tags.filter(t => PROBLEM_TAGS.includes(t)).join(', '),
            tier
          });
        }
      }
      
      // Resolution time
      if (resolvedAt && createdAt) {
        const resHours = (new Date(resolvedAt) - new Date(createdAt)) / (1000 * 60 * 60);
        resolutionTimes.push({ tier, hours: resHours });
        
        // Tier metrics
        if (!tierMetrics[tier]) {
          tierMetrics[tier] = { count: 0, resolutionTimes: [] };
        }
        tierMetrics[tier].count++;
        tierMetrics[tier].resolutionTimes.push(resHours);
      }
    }
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['OPS LIFECYCLE REPORT - Ticket Flow & Resolution Problems', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', 'See Assumptions_Log for methodology', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT A: STATUS DISTRIBUTION ───
    rows.push(['EXHIBIT A: STATUS DISTRIBUTION', '', '', '', '']);
    rows.push(['Status', 'Count', '% of Total', '', '']);
    
    const totalTickets = Object.values(statusCounts).reduce((a, b) => a + b, 0);
    Object.entries(statusCounts).forEach(([status, count]) => {
      const pct = totalTickets > 0 ? ((count / totalTickets) * 100).toFixed(1) + '%' : '0%';
      rows.push([status, count, pct, '', '']);
    });
    rows.push(['TOTAL', totalTickets, '100%', '', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT B: TAG FREQUENCY (Top 20) ───
    rows.push(['EXHIBIT B: TAG FREQUENCY (Top 20)', '', '', '', '']);
    rows.push(['Tag', 'Count', '% of Tickets', '', '']);
    
    const sortedTags = Object.entries(tagFrequency)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 20);
    
    sortedTags.forEach(([tag, count]) => {
      const pct = totalTickets > 0 ? ((count / totalTickets) * 100).toFixed(1) + '%' : '0%';
      rows.push([tag, count, pct, '', '']);
    });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT C: PROBLEM TAGS ───
    rows.push(['EXHIBIT C: PROBLEM TAGS', '', '', '', '']);
    rows.push(['Problem Tag', 'Count', '% of Tickets', 'Status', '']);
    
    Object.entries(problemTagCounts)
      .filter(([_, count]) => count > 0)
      .sort((a, b) => b[1] - a[1])
      .forEach(([tag, count]) => {
        const pct = totalTickets > 0 ? ((count / totalTickets) * 100).toFixed(1) + '%' : '0%';
        rows.push([tag, count, pct, 'Flagged for attention', '']);
      });
    
    if (Object.values(problemTagCounts).every(c => c === 0)) {
      rows.push(['No problem tags found', '', '', '', '']);
    }
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT D: STATUS × PROBLEM TAG MATRIX ───
    rows.push(['EXHIBIT D: STATUS × PROBLEM TAG MATRIX', '', '', '', '']);
    rows.push(['Status', ...PROBLEM_TAGS.slice(0, 4)]);  // First 4 problem tags for readability
    
    ['Open', 'Pending', 'Resolved', 'Closed'].forEach(status => {
      const statusData = statusTagMatrix[status] || {};
      const matrixRow = [status];
      PROBLEM_TAGS.slice(0, 4).forEach(tag => {
        matrixRow.push(statusData[tag] || 0);
      });
      rows.push(matrixRow);
    });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT E: UNRESOLVED TICKET BREAKDOWN ───
    rows.push(['EXHIBIT E: UNRESOLVED TICKETS BY AGE', '', '', '', '']);
    rows.push(['Age Bucket', 'Open', 'Pending', 'Total', 'Risk Level']);
    
    const ageBuckets = [
      { label: '0-24 hours', min: 0, max: 24, risk: 'Normal' },
      { label: '24-48 hours', min: 24, max: 48, risk: 'Warning' },
      { label: '48-72 hours', min: 48, max: 72, risk: 'Elevated' },
      { label: '72+ hours', min: 72, max: Infinity, risk: 'Critical' }
    ];
    
    ageBuckets.forEach(bucket => {
      const openInBucket = ageByStatus.Open.filter(t => t.ageHours >= bucket.min && t.ageHours < bucket.max).length;
      const pendingInBucket = ageByStatus.Pending.filter(t => t.ageHours >= bucket.min && t.ageHours < bucket.max).length;
      rows.push([bucket.label, openInBucket, pendingInBucket, openInBucket + pendingInBucket, bucket.risk]);
    });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT F: RESOLUTION BLOCKERS ───
    rows.push(['EXHIBIT F: RESOLUTION BLOCKERS (Problem Tags + Open > 24hrs)', '', '', '', '']);
    rows.push(['Ticket ID', 'Status', 'Age (Hours)', 'Problem Tags', 'Tier']);
    
    const topBlockers = blockedTickets
      .sort((a, b) => b.ageHours - a.ageHours)
      .slice(0, 20);
    
    if (topBlockers.length > 0) {
      topBlockers.forEach(ticket => {
        rows.push([ticket.ticketId, ticket.status, ticket.ageHours, ticket.problemTags, ticket.tier]);
      });
    } else {
      rows.push(['No blocked tickets found', '', '', '', '']);
    }
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT G: TIER LIFECYCLE COMPARISON ───
    rows.push(['EXHIBIT G: RESOLUTION TIME BY TIER', '', '', '', '']);
    rows.push(['Tier', 'Resolved Count', 'Avg Hours', 'Median Hours', 'P90 Hours']);
    
    Object.entries(tierMetrics).forEach(([tier, metrics]) => {
      const times = metrics.resolutionTimes.sort((a, b) => a - b);
      const avg = times.length > 0 ? (times.reduce((a, b) => a + b, 0) / times.length).toFixed(1) : 'N/A';
      const median = times.length > 0 ? times[Math.floor(times.length / 2)].toFixed(1) : 'N/A';
      const p90 = times.length > 0 ? times[Math.floor(times.length * 0.9)].toFixed(1) : 'N/A';
      rows.push([tier, metrics.count, avg, median, p90]);
    });
    rows.push(['', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '']);
    rows.push(['Data Source:', `${SHEET_NAMES.MVR_RAW_TICKETS} sheet (${totalTickets} tickets)`, '', '', '']);
    rows.push(['Thresholds:', `Open Warning: ${ASSUMPTIONS.PROBLEM_THRESHOLDS.open_hours_warning}hrs, Critical: ${ASSUMPTIONS.PROBLEM_THRESHOLDS.open_hours_critical}hrs`, '', '', '']);
    rows.push(['Problem Tags:', PROBLEM_TAGS.join(', '), '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format the sheet
    formatOpsLifecycleSheet(sheet, normalizedRows.length);
    
    Logger.log(`✅ OPS Lifecycle Report generated with ${normalizedRows.length} rows`);
    Logger.log(`   - ${totalTickets} total tickets analyzed`);
    Logger.log(`   - ${blockedTickets.length} blocked tickets identified`);
    
    return {
      success: true,
      totalTickets,
      blockedTickets: blockedTickets.length,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating OPS lifecycle report: ${e.message}`);
    throw e;
  }
}

/**
 * Format the OPS lifecycle sheet
 */
function formatOpsLifecycleSheet(sheet, totalRows) {
  // Title
  sheet.getRange(1, 1, 1, 5).merge()
    .setBackground('#ea4335').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14);
  
  // Find and format exhibit headers
  for (let i = 1; i <= totalRows; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string' && cellValue.startsWith('EXHIBIT')) {
      sheet.getRange(i, 1, 1, 5)
        .setBackground('#fce8e6').setFontWeight('bold');
    }
    if (cellValue === 'METHODS') {
      sheet.getRange(i, 1, 1, 5)
        .setBackground('#e8f0fe').setFontWeight('bold');
    }
  }
  
  // Auto-resize
  for (let i = 1; i <= 5; i++) {
    sheet.autoResizeColumn(i);
  }
  
  sheet.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// FINANCE RECONCILIATION REPORT - Vendor Billing & Turn ID Lists
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Finance Reconciliation Report
 * Focuses on vendor volumes, Turn ID lists for billing, FREE vs PAID split
 * 
 * @return {Object} Summary of generated report
 */
function generateFinanceReconciliationReport() {
  Logger.log('\n💰 === GENERATING FINANCE RECONCILIATION REPORT ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.FINANCE_AUDIT_TRAIL;
  const reportDate = new Date().toISOString().split('T')[0];
  
  try {
    // Get raw tickets data
    const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      throw new Error('No raw ticket data available. Run initial pull first.');
    }
    
    const data = rawSheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Initialize metrics
    const vendorVolumes = {};  // vendor -> { total, free, paid, turnIds }
    const stateVolumes = {};   // state -> { total, free, paid }
    const monthlyVolumes = {}; // YYYY-MM -> { total, free, paid, vendors: {} }
    const turnIdsByVendor = {};  // vendor -> [turnIds]
    const turnIdsByState = {};   // state -> [turnIds]
    const discrepancies = [];    // Tickets with missing/inconsistent data
    
    let totalChecks = 0;
    let freeChecks = 0;
    let paidChecks = 0;
    
    // Process each ticket
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ticketId = row[colMap['Ticket ID']];
      const turnId = row[colMap['Turn ID']] || '';
      const vendor = row[colMap['State Vendor']] || 'UNKNOWN';
      const state = row[colMap['State']] || '';
      const createdAt = row[colMap['Created At']];
      const status = row[colMap['Status']];
      
      // Only count completed checks (Resolved or Closed)
      const isCompleted = status === 'Resolved' || status === 'Closed';
      if (!isCompleted) continue;
      
      totalChecks++;
      
      // Determine FREE/PAID from STATE_CONFIG
      const stateConfig = STATE_CONFIG[state];
      const isFree = stateConfig ? stateConfig.is_free : false;
      
      if (isFree) {
        freeChecks++;
      } else {
        paidChecks++;
      }
      
      // Vendor volumes
      if (!vendorVolumes[vendor]) {
        vendorVolumes[vendor] = { total: 0, free: 0, paid: 0 };
      }
      vendorVolumes[vendor].total++;
      if (isFree) vendorVolumes[vendor].free++;
      else vendorVolumes[vendor].paid++;
      
      // State volumes
      if (!stateVolumes[state]) {
        stateVolumes[state] = { total: 0, free: 0, paid: 0 };
      }
      stateVolumes[state].total++;
      if (isFree) stateVolumes[state].free++;
      else stateVolumes[state].paid++;
      
      // Monthly volumes
      if (createdAt) {
        const monthKey = new Date(createdAt).toISOString().slice(0, 7); // YYYY-MM
        if (!monthlyVolumes[monthKey]) {
          monthlyVolumes[monthKey] = { total: 0, free: 0, paid: 0, vendors: {} };
        }
        monthlyVolumes[monthKey].total++;
        if (isFree) monthlyVolumes[monthKey].free++;
        else monthlyVolumes[monthKey].paid++;
        
        if (!monthlyVolumes[monthKey].vendors[vendor]) {
          monthlyVolumes[monthKey].vendors[vendor] = 0;
        }
        monthlyVolumes[monthKey].vendors[vendor]++;
      }
      
      // Track Turn IDs (for PAID vendor billing)
      if (turnId && vendor !== 'FREE' && vendor !== 'UNKNOWN') {
        if (!turnIdsByVendor[vendor]) turnIdsByVendor[vendor] = [];
        turnIdsByVendor[vendor].push({ turnId, ticketId, state, date: createdAt });
        
        if (!turnIdsByState[state]) turnIdsByState[state] = [];
        turnIdsByState[state].push({ turnId, ticketId, vendor, date: createdAt });
      }
      
      // Track discrepancies
      if (!turnId && vendor !== 'FREE' && vendor !== 'UNKNOWN') {
        discrepancies.push({
          ticketId,
          issue: 'Missing Turn ID for PAID vendor',
          vendor,
          state
        });
      }
      if (vendor === 'UNKNOWN' && !isFree) {
        discrepancies.push({
          ticketId,
          issue: 'Unknown vendor for PAID state',
          state
        });
      }
    }
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['FINANCE RECONCILIATION REPORT - Vendor Billing & Audit Trail', '', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', 'See Assumptions_Log for state FREE/PAID classification', '', '']);
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT A: SUMMARY TOTALS ───
    rows.push(['EXHIBIT A: COMPLETED CHECKS SUMMARY', '', '', '', '', '']);
    rows.push(['Category', 'Count', '% of Total', '', '', '']);
    rows.push(['Total Completed Checks', totalChecks, '100%', '', '', '']);
    rows.push(['FREE State Checks (DMV Site)', freeChecks, totalChecks > 0 ? ((freeChecks/totalChecks)*100).toFixed(1)+'%' : '0%', '', '', '']);
    rows.push(['PAID Vendor Checks', paidChecks, totalChecks > 0 ? ((paidChecks/totalChecks)*100).toFixed(1)+'%' : '0%', '', '', '']);
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT B: VENDOR BREAKDOWN ───
    rows.push(['EXHIBIT B: VOLUME BY VENDOR', '', '', '', '', '']);
    rows.push(['Vendor', 'Total', 'FREE', 'PAID', '% of Total', '']);
    
    Object.entries(vendorVolumes)
      .sort((a, b) => b[1].total - a[1].total)
      .forEach(([vendor, vol]) => {
        const pct = totalChecks > 0 ? ((vol.total/totalChecks)*100).toFixed(1)+'%' : '0%';
        rows.push([vendor, vol.total, vol.free, vol.paid, pct, '']);
      });
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT C: STATE BREAKDOWN ───
    rows.push(['EXHIBIT C: VOLUME BY STATE (Top 20)', '', '', '', '', '']);
    rows.push(['State', 'Total', 'FREE', 'PAID', '% of Total', 'Vendor Type']);
    
    Object.entries(stateVolumes)
      .sort((a, b) => b[1].total - a[1].total)
      .slice(0, 20)
      .forEach(([state, vol]) => {
        const pct = totalChecks > 0 ? ((vol.total/totalChecks)*100).toFixed(1)+'%' : '0%';
        const vendorType = STATE_CONFIG[state]?.is_free ? 'FREE' : 'PAID';
        rows.push([state, vol.total, vol.free, vol.paid, pct, vendorType]);
      });
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT D: MONTHLY TREND ───
    rows.push(['EXHIBIT D: MONTHLY VOLUME TREND', '', '', '', '', '']);
    rows.push(['Month', 'Total', 'FREE', 'PAID', 'PAID %', '']);
    
    Object.entries(monthlyVolumes)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([month, vol]) => {
        const paidPct = vol.total > 0 ? ((vol.paid/vol.total)*100).toFixed(1)+'%' : '0%';
        rows.push([month, vol.total, vol.free, vol.paid, paidPct, '']);
      });
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT E: BILLABLE TURN IDs BY VENDOR ───
    rows.push(['EXHIBIT E: BILLABLE TURN IDs BY VENDOR (Summary)', '', '', '', '', '']);
    rows.push(['Vendor', 'Billable Count', 'States Covered', 'Date Range', '', '']);
    
    Object.entries(turnIdsByVendor)
      .sort((a, b) => b[1].length - a[1].length)
      .forEach(([vendor, turnIds]) => {
        const states = [...new Set(turnIds.map(t => t.state))].join(', ');
        const dates = turnIds.map(t => t.date).filter(d => d).sort();
        const dateRange = dates.length > 0 ? `${dates[0]?.toString().slice(0,10) || 'N/A'} to ${dates[dates.length-1]?.toString().slice(0,10) || 'N/A'}` : 'N/A';
        rows.push([vendor, turnIds.length, states, dateRange, '', '']);
      });
    rows.push(['', '', '', '', '', '']);
    
    // ─── EXHIBIT F: DISCREPANCIES ───
    rows.push(['EXHIBIT F: DATA DISCREPANCIES', '', '', '', '', '']);
    rows.push(['Ticket ID', 'Issue', 'Vendor', 'State', '', '']);
    
    if (discrepancies.length > 0) {
      discrepancies.slice(0, 50).forEach(d => {
        rows.push([d.ticketId, d.issue, d.vendor || '', d.state, '', '']);
      });
      if (discrepancies.length > 50) {
        rows.push([`... and ${discrepancies.length - 50} more`, '', '', '', '', '']);
      }
    } else {
      rows.push(['No discrepancies found', '', '', '', '', '']);
    }
    rows.push(['', '', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '', '']);
    rows.push(['Data Source:', `${SHEET_NAMES.MVR_RAW_TICKETS} sheet (Resolved/Closed only)`, '', '', '', '']);
    rows.push(['FREE States:', Object.entries(STATE_CONFIG).filter(([k,v]) => v.is_free).map(([k]) => k).join(', '), '', '', '', '']);
    rows.push(['Classification:', 'State → FREE/PAID from STATE_CONFIG; Vendor from ticket State Vendor field', '', '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format the sheet
    formatFinanceSheet(sheet, normalizedRows.length);
    
    // Create separate Turn ID list sheet
    createTurnIdListSheet(ss, turnIdsByVendor, reportDate);
    
    Logger.log(`✅ Finance Reconciliation Report generated`);
    Logger.log(`   - ${totalChecks} completed checks analyzed`);
    Logger.log(`   - ${freeChecks} FREE, ${paidChecks} PAID`);
    Logger.log(`   - ${discrepancies.length} discrepancies found`);
    
    return {
      success: true,
      totalChecks,
      freeChecks,
      paidChecks,
      discrepancies: discrepancies.length,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating Finance report: ${e.message}`);
    throw e;
  }
}

/**
 * Create a dedicated sheet with full Turn ID lists for vendor billing
 */
function createTurnIdListSheet(ss, turnIdsByVendor, reportDate) {
  const sheetName = 'Turn_ID_Billing_List';
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  
  const rows = [
    ['TURN ID BILLING LIST - For Vendor Reconciliation', '', '', '', ''],
    ['Generated:', reportDate, '', '', ''],
    [''],
    ['Vendor', 'Turn ID', 'Ticket ID', 'State', 'Date']
  ];
  
  Object.entries(turnIdsByVendor)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([vendor, turnIds]) => {
      turnIds.forEach(t => {
        const dateStr = t.date ? new Date(t.date).toISOString().split('T')[0] : '';
        rows.push([vendor, t.turnId, t.ticketId, t.state, dateStr]);
      });
    });
  
  sheet.getRange(1, 1, rows.length, 5).setValues(rows);
  
  // Format
  sheet.getRange(1, 1, 1, 5).merge()
    .setBackground('#34a853').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(12);
  sheet.getRange(4, 1, 1, 5).setFontWeight('bold').setBackground('#e6f4ea');
  sheet.setFrozenRows(4);
  
  for (let i = 1; i <= 5; i++) {
    sheet.autoResizeColumn(i);
  }
}

/**
 * Format the Finance Reconciliation sheet
 */
function formatFinanceSheet(sheet, totalRows) {
  // Title
  sheet.getRange(1, 1, 1, 6).merge()
    .setBackground('#34a853').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14);
  
  // Find and format exhibit headers
  for (let i = 1; i <= totalRows; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string' && cellValue.startsWith('EXHIBIT')) {
      sheet.getRange(i, 1, 1, 6)
        .setBackground('#e6f4ea').setFontWeight('bold');
    }
    if (cellValue === 'METHODS') {
      sheet.getRange(i, 1, 1, 6)
        .setBackground('#e8f0fe').setFontWeight('bold');
    }
  }
  
  // Auto-resize
  for (let i = 1; i <= 6; i++) {
    sheet.autoResizeColumn(i);
  }
  
  sheet.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// CEO EFFECTIVENESS REPORT - Program Impact & Detection Metrics
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate CEO Effectiveness Report
 * Focuses on program impact: detection rate, outcomes, coverage, ROI indicators
 * 
 * @return {Object} Summary of generated report
 */
function generateCEOEffectivenessReport() {
  Logger.log('\n📈 === GENERATING CEO EFFECTIVENESS REPORT ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.EXECUTIVE_DASHBOARD;
  const reportDate = new Date().toISOString().split('T')[0];
  
  try {
    // Get raw tickets data
    const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
    if (!rawSheet || rawSheet.getLastRow() <= 1) {
      throw new Error('No raw ticket data available. Run initial pull first.');
    }
    
    const data = rawSheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    // Initialize metrics
    let totalTickets = 0;
    let resolvedTickets = 0;
    let positiveDetections = 0;  // Something found
    let negativeResults = 0;     // Clear/No issues
    let pendingReview = 0;
    
    const outcomesByType = {};   // MVR result type -> count
    const resolutionTimes = [];  // hours to resolve
    const monthlyTrends = {};    // YYYY-MM -> { total, positive, negative }
    const tierCoverage = {};     // tier -> { total, resolved }
    const stateCoverage = {};    // state -> { total, resolved }
    
    // Detection tags (adjust based on actual tag patterns)
    const positiveDetectionTags = ['mvr_positive', 'violation_found', 'action_required', 'suspended', 'revoked', 'expired'];
    const negativeResultTags = ['mvr_clear', 'no_issues', 'valid_license', 'cleared'];
    
    // Process each ticket
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ticketId = row[colMap['Ticket ID']];
      const status = row[colMap['Status']] || '';
      const tagsRaw = row[colMap['Tags']] || '';
      const createdAt = row[colMap['Created At']];
      const resolvedAt = row[colMap['Resolved At']];
      const tier = row[colMap['Tier']] || 'Unknown';
      const state = row[colMap['State']] || 'Unknown';
      const mvrResult = row[colMap['MVR Result']] || row[colMap['Outcome']] || '';
      
      totalTickets++;
      const tags = tagsRaw.split(',').map(t => t.trim().toLowerCase()).filter(t => t);
      
      // Status categorization
      if (status === 'Resolved' || status === 'Closed') {
        resolvedTickets++;
        
        // Resolution time
        if (createdAt && resolvedAt) {
          const resHours = (new Date(resolvedAt) - new Date(createdAt)) / (1000 * 60 * 60);
          resolutionTimes.push(resHours);
        }
        
        // Detection outcome (check tags or result field)
        const hasPositive = tags.some(t => positiveDetectionTags.some(p => t.includes(p))) || 
                           (mvrResult && mvrResult.toLowerCase().match(/positive|violation|suspend|revok|expired/));
        const hasNegative = tags.some(t => negativeResultTags.some(n => t.includes(n))) ||
                           (mvrResult && mvrResult.toLowerCase().match(/clear|valid|no.?issue/));
        
        if (hasPositive) {
          positiveDetections++;
        } else if (hasNegative) {
          negativeResults++;
        }
      } else if (status === 'Pending' || status === 'Open') {
        pendingReview++;
      }
      
      // Outcome types
      if (mvrResult) {
        outcomesByType[mvrResult] = (outcomesByType[mvrResult] || 0) + 1;
      }
      
      // Monthly trends
      if (createdAt) {
        const monthKey = new Date(createdAt).toISOString().slice(0, 7);
        if (!monthlyTrends[monthKey]) {
          monthlyTrends[monthKey] = { total: 0, resolved: 0, positive: 0 };
        }
        monthlyTrends[monthKey].total++;
        if (status === 'Resolved' || status === 'Closed') {
          monthlyTrends[monthKey].resolved++;
        }
      }
      
      // Tier coverage
      if (!tierCoverage[tier]) {
        tierCoverage[tier] = { total: 0, resolved: 0 };
      }
      tierCoverage[tier].total++;
      if (status === 'Resolved' || status === 'Closed') {
        tierCoverage[tier].resolved++;
      }
      
      // State coverage
      if (!stateCoverage[state]) {
        stateCoverage[state] = { total: 0, resolved: 0 };
      }
      stateCoverage[state].total++;
      if (status === 'Resolved' || status === 'Closed') {
        stateCoverage[state].resolved++;
      }
    }
    
    // Calculate key metrics
    const resolutionRate = totalTickets > 0 ? (resolvedTickets / totalTickets) * 100 : 0;
    const detectionRate = resolvedTickets > 0 ? (positiveDetections / resolvedTickets) * 100 : 0;
    const avgResolutionHours = resolutionTimes.length > 0 ? 
      resolutionTimes.reduce((a, b) => a + b, 0) / resolutionTimes.length : 0;
    const medianResolutionHours = resolutionTimes.length > 0 ?
      resolutionTimes.sort((a, b) => a - b)[Math.floor(resolutionTimes.length / 2)] : 0;
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['CEO EFFECTIVENESS REPORT - MVR Program Impact', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', '', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT A: KEY PERFORMANCE INDICATORS ───
    rows.push(['EXHIBIT A: KEY PERFORMANCE INDICATORS', '', '', '', '']);
    rows.push(['KPI', 'Value', 'Target', 'Status', '']);
    rows.push(['Total Checks Processed', totalTickets, '-', 'Baseline', '']);
    rows.push(['Resolution Rate', resolutionRate.toFixed(1) + '%', '90%', resolutionRate >= 90 ? '✓ On Track' : '⚠ Below Target', '']);
    rows.push(['Detection Rate (Positives)', detectionRate.toFixed(1) + '%', 'Monitoring', 'Baseline', '']);
    rows.push(['Avg Resolution Time', avgResolutionHours.toFixed(1) + ' hrs', '< 48 hrs', avgResolutionHours <= 48 ? '✓ On Track' : '⚠ Review', '']);
    rows.push(['Median Resolution Time', medianResolutionHours.toFixed(1) + ' hrs', '< 24 hrs', medianResolutionHours <= 24 ? '✓ On Track' : '⚠ Review', '']);
    rows.push(['Pending Review', pendingReview, '< 50', pendingReview <= 50 ? '✓ OK' : '⚠ High Volume', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT B: DETECTION OUTCOMES ───
    rows.push(['EXHIBIT B: DETECTION OUTCOMES', '', '', '', '']);
    rows.push(['Outcome Category', 'Count', '% of Resolved', '', '']);
    rows.push(['Positive Detections (Issues Found)', positiveDetections, resolvedTickets > 0 ? (positiveDetections/resolvedTickets*100).toFixed(1)+'%' : '0%', '', '']);
    rows.push(['Negative Results (All Clear)', negativeResults, resolvedTickets > 0 ? (negativeResults/resolvedTickets*100).toFixed(1)+'%' : '0%', '', '']);
    rows.push(['Other/Uncategorized', resolvedTickets - positiveDetections - negativeResults, '', '', '']);
    rows.push(['TOTAL RESOLVED', resolvedTickets, '100%', '', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT C: TIER COVERAGE ───
    rows.push(['EXHIBIT C: COVERAGE BY TIER', '', '', '', '']);
    rows.push(['Tier', 'Total Tickets', 'Resolved', 'Resolution %', 'Cadence']);
    
    Object.entries(tierCoverage)
      .sort((a, b) => a[0].localeCompare(b[0]))
      .forEach(([tier, metrics]) => {
        const resPct = metrics.total > 0 ? (metrics.resolved / metrics.total * 100).toFixed(1) + '%' : '0%';
        const tierConfig = TIER_CONFIG[tier] || { cadence_days: '-', description: '-' };
        rows.push([tier, metrics.total, metrics.resolved, resPct, tierConfig.description || '-']);
      });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT D: MONTHLY TREND ───
    rows.push(['EXHIBIT D: MONTHLY PERFORMANCE TREND', '', '', '', '']);
    rows.push(['Month', 'Total', 'Resolved', 'Resolution %', 'Trend']);
    
    const sortedMonths = Object.entries(monthlyTrends).sort((a, b) => a[0].localeCompare(b[0]));
    let prevResRate = null;
    sortedMonths.forEach(([month, metrics]) => {
      const resRate = metrics.total > 0 ? (metrics.resolved / metrics.total * 100) : 0;
      const trend = prevResRate !== null ? (resRate > prevResRate ? '↑' : resRate < prevResRate ? '↓' : '→') : '-';
      rows.push([month, metrics.total, metrics.resolved, resRate.toFixed(1) + '%', trend]);
      prevResRate = resRate;
    });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT E: STATE COVERAGE (Top 15) ───
    rows.push(['EXHIBIT E: STATE COVERAGE (Top 15)', '', '', '', '']);
    rows.push(['State', 'Total', 'Resolved', 'Resolution %', 'Tier']);
    
    Object.entries(stateCoverage)
      .sort((a, b) => b[1].total - a[1].total)
      .slice(0, 15)
      .forEach(([state, metrics]) => {
        const resPct = metrics.total > 0 ? (metrics.resolved / metrics.total * 100).toFixed(1) + '%' : '0%';
        const tier = STATE_CONFIG[state]?.tier || '-';
        rows.push([state, metrics.total, metrics.resolved, resPct, tier]);
      });
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT F: PROGRAM VALUE INDICATORS ───
    rows.push(['EXHIBIT F: PROGRAM VALUE INDICATORS', '', '', '', '']);
    rows.push(['Metric', 'Value', 'Interpretation', '', '']);
    rows.push(['Proactive Detections', positiveDetections, 'Issues caught before incident', '', '']);
    rows.push(['States Covered', Object.keys(stateCoverage).length, 'Geographic reach', '', '']);
    rows.push(['Avg Turnaround', avgResolutionHours.toFixed(1) + ' hrs', 'Operational efficiency', '', '']);
    
    // Cost avoidance estimate (rough)
    const costPerIncident = 50000;  // Industry avg for driving-related incident
    const estimatedSavings = positiveDetections * costPerIncident * 0.1;  // Conservative 10% would have caused incident
    rows.push(['Est. Risk Mitigation Value', '$' + estimatedSavings.toLocaleString(), 'Conservative @ 10% incident prevention', '', '']);
    rows.push(['', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '']);
    rows.push(['Data Source:', `${SHEET_NAMES.MVR_RAW_TICKETS} sheet`, '', '', '']);
    rows.push(['Detection Logic:', 'Positive = tags/results indicating violations, suspensions, expirations', '', '', '']);
    rows.push(['Cost Model:', '$50,000 avg incident cost × 10% prevention rate (conservative)', '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format the sheet
    formatCEOSheet(sheet, normalizedRows.length);
    
    Logger.log(`✅ CEO Effectiveness Report generated`);
    Logger.log(`   - ${totalTickets} total tickets, ${resolvedTickets} resolved`);
    Logger.log(`   - Detection rate: ${detectionRate.toFixed(1)}%`);
    Logger.log(`   - Avg resolution: ${avgResolutionHours.toFixed(1)} hours`);
    
    return {
      success: true,
      totalTickets,
      resolvedTickets,
      detectionRate: detectionRate.toFixed(1),
      avgResolutionHours: avgResolutionHours.toFixed(1),
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating CEO report: ${e.message}`);
    throw e;
  }
}

/**
 * Format the CEO Effectiveness sheet
 */
function formatCEOSheet(sheet, totalRows) {
  // Title
  sheet.getRange(1, 1, 1, 5).merge()
    .setBackground('#4285f4').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14);
  
  // Find and format exhibit headers and KPI rows
  for (let i = 1; i <= totalRows; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string') {
      if (cellValue.startsWith('EXHIBIT')) {
        sheet.getRange(i, 1, 1, 5)
          .setBackground('#e8f0fe').setFontWeight('bold');
      }
      if (cellValue === 'METHODS') {
        sheet.getRange(i, 1, 1, 5)
          .setBackground('#f8f9fa').setFontWeight('bold');
      }
    }
    // Highlight status column for KPIs
    const statusValue = sheet.getRange(i, 4).getValue();
    if (typeof statusValue === 'string') {
      if (statusValue.includes('✓')) {
        sheet.getRange(i, 4).setBackground('#e6f4ea').setFontColor('#137333');
      } else if (statusValue.includes('⚠')) {
        sheet.getRange(i, 4).setBackground('#fce8e6').setFontColor('#c5221f');
      }
    }
  }
  
  // Auto-resize
  for (let i = 1; i <= 5; i++) {
    sheet.autoResizeColumn(i);
  }
  
  sheet.setFrozenRows(2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ANNUAL PROJECTION REPORT - Capacity Planning with Scenarios
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Annual Projection Report
 * Projects annual check volumes, FTE needs, and capacity by Low/Mid/High scenarios
 * 
 * @param {number} totalProfiles - Total employee profiles to project (or uses mid estimate)
 * @return {Object} Summary of projections
 */
function generateAnnualProjection(totalProfiles) {
  Logger.log('\n📅 === GENERATING ANNUAL PROJECTION REPORT ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.ANNUAL_PROJECTION;
  const reportDate = new Date().toISOString().split('T')[0];
  const projectionYear = new Date().getFullYear() + 1;  // Project next year
  
  try {
    // If no total profiles provided, try to estimate from data or use placeholder
    const estimatedProfiles = totalProfiles || 10000;  // Placeholder - adjust based on actual data
    
    // Get assumptions
    const enrolledPct = ASSUMPTIONS.ENROLLED_PCT;
    const activePct = ASSUMPTIONS.ACTIVE_PROFILE_PCT;
    const freeStatePct = ASSUMPTIONS.FREE_STATE_PCT;
    const capacity = CAPACITY_CONFIG;
    
    // Calculate scenarios
    const scenarios = ['low', 'mid', 'high'];
    const projections = {};
    
    scenarios.forEach(scenario => {
      // How many profiles are actively being monitored?
      const enrolledProfiles = estimatedProfiles * (enrolledPct[scenario] / 100);
      const activeProfiles = enrolledProfiles * (activePct[scenario] / 100);
      
      // Calculate annual checks based on tier distribution
      // Simplified: assume even distribution across tiers (adjust with actual data)
      const checksPerProfile = {
        tier1: TIER_CONFIG['Tier 1']?.checks_per_year || 6,
        tier2: TIER_CONFIG['Tier 2']?.checks_per_year || 4,
        tier3: TIER_CONFIG['Tier 3']?.checks_per_year || 2
      };
      
      // Assume distribution: 10% T1, 20% T2, 70% T3 (adjust based on actual)
      const tierDist = { tier1: 0.10, tier2: 0.20, tier3: 0.70 };
      
      const weightedChecksPerProfile = 
        (tierDist.tier1 * checksPerProfile.tier1) +
        (tierDist.tier2 * checksPerProfile.tier2) +
        (tierDist.tier3 * checksPerProfile.tier3);
      
      const totalAnnualChecks = Math.round(activeProfiles * weightedChecksPerProfile);
      
      // Monthly distribution (assume even)
      const checksPerMonth = Math.round(totalAnnualChecks / 12);
      
      // Working days per month
      const workingDaysPerMonth = capacity.WORKING_DAYS_PER_MONTH;
      const checksPerDay = Math.round(checksPerMonth / workingDaysPerMonth);
      
      // FTE calculation
      const checksPerAgentDay = capacity.EFFECTIVE_CHECKS_PER_DAY;
      const fteNeeded = checksPerDay / checksPerAgentDay;
      
      // Cost split
      const freeStateChecks = Math.round(totalAnnualChecks * (freeStatePct[scenario] / 100));
      const paidChecks = totalAnnualChecks - freeStateChecks;
      
      projections[scenario] = {
        enrolledProfiles: Math.round(enrolledProfiles),
        activeProfiles: Math.round(activeProfiles),
        totalAnnualChecks,
        checksPerMonth,
        checksPerDay,
        fteNeeded: fteNeeded.toFixed(2),
        freeStateChecks,
        paidChecks
      };
    });
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['ANNUAL PROJECTION REPORT - ' + projectionYear, '', '', '', '']);
    rows.push(['Generated:', reportDate, '', 'Base Profiles:', estimatedProfiles.toLocaleString()]);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT A: INPUT ASSUMPTIONS ───
    rows.push(['EXHIBIT A: INPUT ASSUMPTIONS', '', '', '', '']);
    rows.push(['Parameter', 'Low', 'Mid', 'High', 'Source']);
    rows.push(['Enrolled %', enrolledPct.low + '%', enrolledPct.mid + '%', enrolledPct.high + '%', 'ASSUMPTIONS.ENROLLED_PCT']);
    rows.push(['Active Profile %', activePct.low + '%', activePct.mid + '%', activePct.high + '%', 'ASSUMPTIONS.ACTIVE_PROFILE_PCT']);
    rows.push(['FREE State %', freeStatePct.low + '%', freeStatePct.mid + '%', freeStatePct.high + '%', 'ASSUMPTIONS.FREE_STATE_PCT']);
    rows.push(['MVR Time Allocation', capacity.MVR_TIME_ALLOCATION_PCT + '%', '', '', 'CAPACITY_CONFIG']);
    rows.push(['Checks/Agent/Day (@ 10%)', capacity.EFFECTIVE_CHECKS_PER_DAY, '', '', 'CAPACITY_CONFIG']);
    rows.push(['Working Days/Month', capacity.WORKING_DAYS_PER_MONTH, '', '', 'CAPACITY_CONFIG']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT B: VOLUME PROJECTIONS ───
    rows.push(['EXHIBIT B: ANNUAL VOLUME PROJECTIONS', '', '', '', '']);
    rows.push(['Metric', 'Low', 'Mid', 'High', 'Notes']);
    rows.push(['Enrolled Profiles', projections.low.enrolledProfiles.toLocaleString(), projections.mid.enrolledProfiles.toLocaleString(), projections.high.enrolledProfiles.toLocaleString(), '']);
    rows.push(['Active Profiles', projections.low.activeProfiles.toLocaleString(), projections.mid.activeProfiles.toLocaleString(), projections.high.activeProfiles.toLocaleString(), '']);
    rows.push(['Total Annual Checks', projections.low.totalAnnualChecks.toLocaleString(), projections.mid.totalAnnualChecks.toLocaleString(), projections.high.totalAnnualChecks.toLocaleString(), '']);
    rows.push(['Avg Checks/Month', projections.low.checksPerMonth.toLocaleString(), projections.mid.checksPerMonth.toLocaleString(), projections.high.checksPerMonth.toLocaleString(), '']);
    rows.push(['Avg Checks/Day', projections.low.checksPerDay, projections.mid.checksPerDay, projections.high.checksPerDay, 'Working days only']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT C: FTE REQUIREMENTS ───
    rows.push(['EXHIBIT C: FTE REQUIREMENTS (@ 10% MVR Time)', '', '', '', '']);
    rows.push(['Scenario', 'FTE Needed', 'Current FTE', 'Gap', 'Recommendation']);
    const currentFTE = 1;  // Adjust based on actual
    rows.push(['Low', projections.low.fteNeeded, currentFTE, (projections.low.fteNeeded - currentFTE).toFixed(2), parseFloat(projections.low.fteNeeded) <= currentFTE ? 'Sufficient' : 'May need support']);
    rows.push(['Mid', projections.mid.fteNeeded, currentFTE, (projections.mid.fteNeeded - currentFTE).toFixed(2), parseFloat(projections.mid.fteNeeded) <= currentFTE ? 'Sufficient' : 'Plan for hiring']);
    rows.push(['High', projections.high.fteNeeded, currentFTE, (projections.high.fteNeeded - currentFTE).toFixed(2), parseFloat(projections.high.fteNeeded) <= currentFTE ? 'Sufficient' : 'Priority hiring']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT D: COST BREAKDOWN ───
    rows.push(['EXHIBIT D: VENDOR COST PROJECTION', '', '', '', '']);
    rows.push(['Category', 'Low', 'Mid', 'High', 'Notes']);
    rows.push(['FREE State Checks', projections.low.freeStateChecks.toLocaleString(), projections.mid.freeStateChecks.toLocaleString(), projections.high.freeStateChecks.toLocaleString(), 'DMV Site - No vendor cost']);
    rows.push(['PAID Vendor Checks', projections.low.paidChecks.toLocaleString(), projections.mid.paidChecks.toLocaleString(), projections.high.paidChecks.toLocaleString(), 'Vendor billing applies']);
    rows.push(['', '', '', '', '']);
    
    // ─── EXHIBIT E: MONTHLY CAPACITY MATRIX ───
    rows.push(['EXHIBIT E: MONTHLY CAPACITY CHECK (Mid Scenario)', '', '', '', '']);
    rows.push(['Month', 'Projected Checks', 'Available Capacity', 'Utilization %', 'Status']);
    
    const midMonthlyChecks = projections.mid.checksPerMonth;
    const monthlyCapacity = currentFTE * capacity.EFFECTIVE_CHECKS_PER_DAY * capacity.WORKING_DAYS_PER_MONTH;
    
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    months.forEach(month => {
      const utilization = (midMonthlyChecks / monthlyCapacity * 100);
      const status = utilization <= 80 ? '✓ OK' : utilization <= 100 ? '⚠ Tight' : '❌ Over';
      rows.push([month, midMonthlyChecks, monthlyCapacity, utilization.toFixed(0) + '%', status]);
    });
    rows.push(['', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '']);
    rows.push(['Tier Distribution:', '10% T1 (6x/yr), 20% T2 (4x/yr), 70% T3 (2x/yr)', '', '', '']);
    rows.push(['Capacity Model:', `${capacity.MAX_CHECKS_PER_DAY} max/day @ 100%, ${capacity.EFFECTIVE_CHECKS_PER_DAY}/day @ ${capacity.MVR_TIME_ALLOCATION_PCT}%`, '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format
    formatProjectionSheet(sheet, normalizedRows.length, '#fbbc04');  // Yellow theme
    
    Logger.log(`✅ Annual Projection generated for ${projectionYear}`);
    Logger.log(`   - Mid scenario: ${projections.mid.totalAnnualChecks.toLocaleString()} annual checks`);
    Logger.log(`   - FTE needed (mid): ${projections.mid.fteNeeded}`);
    
    return {
      success: true,
      year: projectionYear,
      projections,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating annual projection: ${e.message}`);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// OPERATIVE PLAN - Monthly Wave Calendar
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Operative Plan
 * Monthly wave calendar showing check schedule by tier and state
 * 
 * @param {number} year - Year to plan (defaults to next year)
 * @return {Object} Summary of operative plan
 */
function generateOperativePlan(year) {
  Logger.log('\n📋 === GENERATING OPERATIVE PLAN ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.OPERATIVE_PLAN;
  const reportDate = new Date().toISOString().split('T')[0];
  const planYear = year || (new Date().getFullYear() + 1);
  
  try {
    // Build wave calendar
    // Tier 1: Every 60 days = 6x/year (Jan, Mar, May, Jul, Sep, Nov)
    // Tier 2: Every 90 days = 4x/year (Jan, Apr, Jul, Oct)  
    // Tier 3: Every 180 days = 2x/year (Jan, Jul)
    
    const waveSchedule = {
      'Tier 1': [1, 3, 5, 7, 9, 11],      // Jan, Mar, May, Jul, Sep, Nov
      'Tier 2': [1, 4, 7, 10],             // Jan, Apr, Jul, Oct
      'Tier 3': [1, 7]                      // Jan, Jul
    };
    
    const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    
    // Count states by tier
    const statesByTier = { 'Tier 1': [], 'Tier 2': [], 'Tier 3': [] };
    Object.entries(STATE_CONFIG).forEach(([code, config]) => {
      if (statesByTier[config.tier]) {
        statesByTier[config.tier].push(code);
      }
    });
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['OPERATIVE PLAN - ' + planYear + ' Wave Calendar', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // ─── EXHIBIT A: TIER OVERVIEW ───
    rows.push(['EXHIBIT A: TIER OVERVIEW', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Tier', 'Cadence', 'Checks/Year', 'States', 'Wave Months', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 1', '60 days', '6', statesByTier['Tier 1'].length, 'Jan, Mar, May, Jul, Sep, Nov', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 2', '90 days', '4', statesByTier['Tier 2'].length, 'Jan, Apr, Jul, Oct', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 3', '180 days', '2', statesByTier['Tier 3'].length, 'Jan, Jul', '', '', '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // ─── EXHIBIT B: MONTHLY WAVE CALENDAR ───
    rows.push(['EXHIBIT B: MONTHLY WAVE CALENDAR', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // Header row with months
    const calendarHeader = ['Tier', ...monthNames];
    rows.push(calendarHeader);
    
    // Wave schedule by tier
    Object.entries(waveSchedule).forEach(([tier, months]) => {
      const calendarRow = [tier];
      for (let m = 1; m <= 12; m++) {
        calendarRow.push(months.includes(m) ? '🔵 WAVE' : '-');
      }
      rows.push(calendarRow);
    });
    rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // ─── EXHIBIT C: ESTIMATED MONTHLY WORKLOAD ───
    rows.push(['EXHIBIT C: MONTHLY WORKLOAD DISTRIBUTION', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Month', 'T1 States', 'T2 States', 'T3 States', 'Total States Active', 'Relative Load', '', '', '', '', '', '', '']);
    
    for (let m = 1; m <= 12; m++) {
      const t1Active = waveSchedule['Tier 1'].includes(m) ? statesByTier['Tier 1'].length : 0;
      const t2Active = waveSchedule['Tier 2'].includes(m) ? statesByTier['Tier 2'].length : 0;
      const t3Active = waveSchedule['Tier 3'].includes(m) ? statesByTier['Tier 3'].length : 0;
      const totalActive = t1Active + t2Active + t3Active;
      const load = totalActive > 40 ? 'Heavy' : totalActive > 20 ? 'Medium' : 'Light';
      rows.push([monthNames[m-1], t1Active, t2Active, t3Active, totalActive, load, '', '', '', '', '', '', '']);
    }
    rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // ─── EXHIBIT D: STATES BY TIER ───
    rows.push(['EXHIBIT D: STATES BY TIER', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 1 States:', statesByTier['Tier 1'].join(', '), '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 2 States:', statesByTier['Tier 2'].join(', '), '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Tier 3 States:', statesByTier['Tier 3'].join(', ').substring(0, 100) + '...', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '', '', '', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Wave Logic:', 'Tier 1=60d cadence, Tier 2=90d cadence, Tier 3=180d cadence', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['State Assignment:', 'From STATE_CONFIG in Config.gs', '', '', '', '', '', '', '', '', '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '', '', '', '', '', '', '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format
    formatProjectionSheet(sheet, normalizedRows.length, '#673ab7');  // Purple theme
    
    // Highlight wave cells
    for (let r = 1; r <= normalizedRows.length; r++) {
      for (let c = 2; c <= 13; c++) {
        const val = sheet.getRange(r, c).getValue();
        if (val === '🔵 WAVE') {
          sheet.getRange(r, c).setBackground('#e1bee7');
        }
      }
    }
    
    Logger.log(`✅ Operative Plan generated for ${planYear}`);
    
    return {
      success: true,
      year: planYear,
      statesByTier,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating operative plan: ${e.message}`);
    throw e;
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// GROWTH PROJECTION - 5-Year Table
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Growth Projection Report
 * 5-year projection with hiring triggers based on growth scenarios
 * 
 * @param {number} baseProfiles - Starting profile count (or estimate)
 * @return {Object} Summary of growth projection
 */
function generateGrowthProjection(baseProfiles) {
  Logger.log('\n📈 === GENERATING GROWTH PROJECTION REPORT ===\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = SHEET_NAMES.GROWTH_PROJECTION;
  const reportDate = new Date().toISOString().split('T')[0];
  const startYear = new Date().getFullYear();
  
  try {
    const base = baseProfiles || 10000;  // Placeholder
    const growth = ASSUMPTIONS.GROWTH_SCENARIOS;
    const capacity = CAPACITY_CONFIG;
    
    // Project 5 years
    const years = [startYear, startYear+1, startYear+2, startYear+3, startYear+4];
    const projections = {
      conservative: [],
      moderate: [],
      aggressive: []
    };
    
    // Calculate projections for each scenario
    const scenarioGrowth = {
      conservative: growth.conservative / 100,
      moderate: growth.moderate / 100,
      aggressive: growth.aggressive / 100
    };
    
    Object.keys(projections).forEach(scenario => {
      let profiles = base;
      let currentFTE = 1;
      
      years.forEach((year, idx) => {
        if (idx > 0) {
          profiles = Math.round(profiles * (1 + scenarioGrowth[scenario]));
        }
        
        // Estimate annual checks (using mid assumptions)
        const activeProfiles = profiles * (ASSUMPTIONS.ENROLLED_PCT.mid / 100) * (ASSUMPTIONS.ACTIVE_PROFILE_PCT.mid / 100);
        const avgChecksPerProfile = 2.6;  // Weighted avg
        const annualChecks = Math.round(activeProfiles * avgChecksPerProfile);
        const monthlyChecks = Math.round(annualChecks / 12);
        const dailyChecks = Math.round(monthlyChecks / capacity.WORKING_DAYS_PER_MONTH);
        
        // FTE needed
        const fteNeeded = dailyChecks / capacity.EFFECTIVE_CHECKS_PER_DAY;
        const fteGap = fteNeeded - currentFTE;
        
        // Hiring trigger
        let hiringAction = 'Maintain';
        if (fteGap >= 1) {
          hiringAction = `Hire ${Math.ceil(fteGap)} FTE`;
          currentFTE += Math.ceil(fteGap);
        } else if (fteGap >= 0.5) {
          hiringAction = 'Consider part-time/support';
        }
        
        projections[scenario].push({
          year,
          profiles,
          activeProfiles: Math.round(activeProfiles),
          annualChecks,
          dailyChecks,
          fteNeeded: fteNeeded.toFixed(2),
          currentFTE,
          fteGap: fteGap.toFixed(2),
          hiringAction
        });
      });
    });
    
    // Get or create output sheet
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    // Build report rows
    const rows = [];
    
    // ─── HEADER ───
    rows.push(['5-YEAR GROWTH PROJECTION', '', '', '', '', '', '']);
    rows.push(['Generated:', reportDate, '', 'Base Year:', startYear, 'Base Profiles:', base.toLocaleString()]);
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── EXHIBIT A: GROWTH ASSUMPTIONS ───
    rows.push(['EXHIBIT A: GROWTH SCENARIOS', '', '', '', '', '', '']);
    rows.push(['Scenario', 'Annual Growth', 'Description', '', '', '', '']);
    rows.push(['Conservative', growth.conservative + '%', 'Minimal expansion, stable client base', '', '', '', '']);
    rows.push(['Moderate', growth.moderate + '%', 'Steady growth, new client acquisition', '', '', '', '']);
    rows.push(['Aggressive', growth.aggressive + '%', 'Rapid expansion, major account wins', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── EXHIBIT B: CONSERVATIVE SCENARIO ───
    rows.push(['EXHIBIT B: CONSERVATIVE SCENARIO (' + growth.conservative + '% Annual Growth)', '', '', '', '', '', '']);
    rows.push(['Year', 'Profiles', 'Active', 'Annual Checks', 'Daily Load', 'FTE Needed', 'Hiring Action']);
    projections.conservative.forEach(p => {
      rows.push([p.year, p.profiles.toLocaleString(), p.activeProfiles.toLocaleString(), p.annualChecks.toLocaleString(), p.dailyChecks, p.fteNeeded, p.hiringAction]);
    });
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── EXHIBIT C: MODERATE SCENARIO ───
    rows.push(['EXHIBIT C: MODERATE SCENARIO (' + growth.moderate + '% Annual Growth)', '', '', '', '', '', '']);
    rows.push(['Year', 'Profiles', 'Active', 'Annual Checks', 'Daily Load', 'FTE Needed', 'Hiring Action']);
    projections.moderate.forEach(p => {
      rows.push([p.year, p.profiles.toLocaleString(), p.activeProfiles.toLocaleString(), p.annualChecks.toLocaleString(), p.dailyChecks, p.fteNeeded, p.hiringAction]);
    });
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── EXHIBIT D: AGGRESSIVE SCENARIO ───
    rows.push(['EXHIBIT D: AGGRESSIVE SCENARIO (' + growth.aggressive + '% Annual Growth)', '', '', '', '', '', '']);
    rows.push(['Year', 'Profiles', 'Active', 'Annual Checks', 'Daily Load', 'FTE Needed', 'Hiring Action']);
    projections.aggressive.forEach(p => {
      rows.push([p.year, p.profiles.toLocaleString(), p.activeProfiles.toLocaleString(), p.annualChecks.toLocaleString(), p.dailyChecks, p.fteNeeded, p.hiringAction]);
    });
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── EXHIBIT E: HIRING TRIGGERS SUMMARY ───
    rows.push(['EXHIBIT E: HIRING TRIGGER SUMMARY', '', '', '', '', '', '']);
    rows.push(['Threshold', 'Value', 'Action', '', '', '', '']);
    rows.push(['FTE Gap ≥ 1.0', 'Immediate', 'Hire full-time equivalent', '', '', '', '']);
    rows.push(['FTE Gap 0.5-1.0', 'Plan', 'Consider part-time/contractor support', '', '', '', '']);
    rows.push(['FTE Gap < 0.5', 'Monitor', 'Current capacity sufficient', '', '', '', '']);
    rows.push(['', '', '', '', '', '', '']);
    
    // ─── METHODS ───
    rows.push(['METHODS', '', '', '', '', '', '']);
    rows.push(['Profile Conversion:', `${ASSUMPTIONS.ENROLLED_PCT.mid}% enrolled × ${ASSUMPTIONS.ACTIVE_PROFILE_PCT.mid}% active`, '', '', '', '', '']);
    rows.push(['Checks/Profile:', '2.6 avg (weighted by tier distribution)', '', '', '', '', '']);
    rows.push(['Capacity:', `${capacity.EFFECTIVE_CHECKS_PER_DAY}/day/FTE @ ${capacity.MVR_TIME_ALLOCATION_PCT}% MVR time`, '', '', '', '', '']);
    rows.push(['Full assumptions:', `See ${SHEET_NAMES.ASSUMPTIONS_LOG} sheet`, '', '', '', '', '']);
    
    // Write all rows
    const maxCols = Math.max(...rows.map(r => r.length));
    const normalizedRows = rows.map(r => {
      while (r.length < maxCols) r.push('');
      return r;
    });
    
    sheet.getRange(1, 1, normalizedRows.length, maxCols).setValues(normalizedRows);
    
    // Format
    formatProjectionSheet(sheet, normalizedRows.length, '#0097a7');  // Teal theme
    
    Logger.log(`✅ Growth Projection generated for ${startYear}-${startYear+4}`);
    
    return {
      success: true,
      years: `${startYear}-${startYear+4}`,
      projections,
      generated: reportDate
    };
    
  } catch (e) {
    Logger.log(`❌ Error generating growth projection: ${e.message}`);
    throw e;
  }
}

/**
 * Shared formatting function for projection sheets
 */
function formatProjectionSheet(sheet, totalRows, themeColor) {
  // Title
  sheet.getRange(1, 1, 1, 7).merge()
    .setBackground(themeColor).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(14);
  
  // Find and format exhibit headers
  for (let i = 1; i <= totalRows; i++) {
    const cellValue = sheet.getRange(i, 1).getValue();
    if (typeof cellValue === 'string') {
      if (cellValue.startsWith('EXHIBIT')) {
        const lightColor = themeColor + '20';  // Add transparency
        sheet.getRange(i, 1, 1, 7)
          .setBackground('#f5f5f5').setFontWeight('bold');
      }
      if (cellValue === 'METHODS') {
        sheet.getRange(i, 1, 1, 7)
          .setBackground('#f8f9fa').setFontWeight('bold');
      }
    }
  }
  
  // Auto-resize columns
  const lastCol = sheet.getLastColumn() || 7;
  for (let i = 1; i <= lastCol; i++) {
    sheet.autoResizeColumn(i);
  }
  
  sheet.setFrozenRows(2);
}
