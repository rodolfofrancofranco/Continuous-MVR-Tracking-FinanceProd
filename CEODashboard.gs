/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * CEO EXECUTIVE DASHBOARD - High-Level Operational Assessment
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * Professional executive dashboard with:
 * - KPI cards with big numbers and trend indicators
 * - Sparklines for visual trends
 * - Charts for operational effectiveness
 * - Summary insights for quick decision-making
 * 
 * Designed for CEO-level operational efficiency and effectiveness assessment
 */

// ═══════════════════════════════════════════════════════════════════════════════
// CEO DASHBOARD CREATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create the CEO Executive Dashboard with live formulas and professional styling
 */
function createCEODashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Get or create sheet
  let sheet = ss.getSheetByName('CEO Dashboard');
  if (sheet) {
    const response = ui.alert(
      '⚠️ CEO Dashboard Exists',
      'This will recreate the dashboard. Proceed?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }
  
  sheet = ss.insertSheet('CEO Dashboard');
  
  // Professional setup
  sheet.setHiddenGridlines(true);
  
  // Column widths for 4-card layout
  const colWidths = [20, 160, 160, 160, 160, 20, 220, 120, 120, 120];
  colWidths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  
  let row = 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // Precompute aggregations using Finance standards (do not modify Finance dashboard)
  const agg = computeCEOMetrics({ days: 30, months: 6 });

  // Load raw history for direct dashboard calculations and views (row-level operations)
  const _history = loadHistoryData();
  const _hmap = _history.headerMap;
  const _rawData = _history.data;

  // Helper: numeric KPI values will use config.valueNumber when present
  function kpiConfigFromNumber(n, label, opts) {
    opts = opts || {};
    return Object.assign({ valueNumber: n, label: label }, opts);
  }

  // HEADER SECTION
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 4).merge()
    .setValue('📈 EXECUTIVE DASHBOARD')
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontSize(22)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(row, 55);
  row++;
  
  // Subtitle with dynamic date
  sheet.getRange(row, 2, 1, 4).merge()
    .setFormula('="MVR Operations Performance | As of "&TEXT(NOW(),"mmmm d, yyyy • h:mm AM/PM")')
    .setFontColor(COLORS.SLATE)
    .setFontSize(11)
    .setHorizontalAlignment('center');
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // KPI CARDS ROW 1 - Volume & Resolution
  // ═══════════════════════════════════════════════════════════════════════════
  const kpi1Row = row;
  
  // KPI 1: Total Tickets (30 days) - computed from Last Updated per Finance standards
  createExecutiveKPI(sheet, row, 2, kpiConfigFromNumber(agg.totalTickets30, 'Tickets (30 Days)', { trend: agg.ticketsTrend, format: '#,##0', sparkline: 'column', color: COLORS.ROYAL_BLUE }));
  
  // KPI 2: SLA Compliance Rate (computed)
  createExecutiveKPI(sheet, row, 3, kpiConfigFromNumber(agg.slaCompliance, 'SLA Compliance', { format: '0.0%', color: COLORS.EMERALD, isPercent: true }));
  
  // KPI 3: Detection Rate (computed)
  createExecutiveKPI(sheet, row, 4, kpiConfigFromNumber(agg.detectionRate, 'Detection Rate', { format: '0.0%', color: COLORS.AMBER, isPercent: true }));
  
  // KPI 4: Avg Resolution Time (computed)
  createExecutiveKPI(sheet, row, 5, kpiConfigFromNumber(agg.avgResolutionHours, 'Avg Resolution', { format: '0.0" hrs"', color: COLORS.COST_HEADER }));
  
  row += 6;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // KPI CARDS ROW 2 - Partners & Vendors
  // ═══════════════════════════════════════════════════════════════════════════
  
  // KPI 5: Active Partners (row-level unique partners across rows)
  // compute unique partners from raw rows in the window
  const partnerIdx = _hmap['Partner Name'];
  const partnerSet = {};
  const cutoffForPartners = new Date((new Date()).getTime() - (30 * 24 * 3600 * 1000));
  for (let i = 0; i < _rawData.length; i++) {
    const r = _rawData[i];
    const last = safeParseDate(r[_hmap['Last Updated']] || r[_hmap['Date Created']]);
    if (!last || last < cutoffForPartners) continue;
    const p = (r[partnerIdx] || '').toString().trim();
    if (p) partnerSet[p] = true;
  }
  const activePartnersCount = Object.keys(partnerSet).length;
  createExecutiveKPI(sheet, row, 2, kpiConfigFromNumber(activePartnersCount, 'Active Partners', { format: '#,##0', color: COLORS.CAPACITY_HEADER }));
  
  // KPI 6: Billable Checks (30 days) - row-level (each processed row counts)
  createExecutiveKPI(sheet, row, 3, kpiConfigFromNumber(agg.billable30, 'Billable (30 Days)', { format: '#,##0', color: COLORS.EMERALD }));
  
  // KPI 7: Cost per Check (based on vendor, not Is Billable which doesn't exist)
  // KPI 7: Billable % (30 days) — operational focus; computed from aggregations
  createExecutiveKPI(sheet, row, 4, kpiConfigFromNumber( agg.totalTickets30 ? (agg.billable30 / agg.totalTickets30) : 0, 'Billable % (30 Days)', { format: '0.0%', color: COLORS.COST_HEADER, isPercent: true } ));
  
  // KPI 8: Open Tickets (Current) - row-level count
  const statusIdxForKPI = _hmap['Status'];
  let openPendingCount = 0;
  for (let i = 0; i < _rawData.length; i++) {
    const s = ('' + (_rawData[i][statusIdxForKPI] || '')).toString().trim();
    if (s === 'Open' || s === 'Pending') openPendingCount++;
  }
  createExecutiveKPI(sheet, row, 5, kpiConfigFromNumber(openPendingCount, 'Open/Pending Now', { format: '#,##0', color: COLORS.RED, isAlert: true }));
  
  row += 6;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION: STATUS DISTRIBUTION
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 4).merge()
    .setValue('📊 OPERATIONAL STATUS')
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  // Status headers
  const statusHeaders = ['Status', 'Count', '% of Total', 'Trend'];
  sheet.getRange(row, 2, 1, 4).setValues([statusHeaders])
    .setBackground(COLORS.MEDIUM_GRAY)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  row++;
  
  const statuses = ['Open', 'Pending', 'Resolved', 'Closed'];
  const statusColors = {
    'Open': '#fef3c7',
    'Pending': '#dbeafe',
    'Resolved': '#d1fae5',
    'Closed': '#e5e7eb'
  };
  const statusDataStart = row;
  
  statuses.forEach((status, idx) => {
    const bgColor = statusColors[status];
    
    sheet.getRange(row, 2).setValue(status)
      .setBackground(bgColor).setFontWeight('bold');
    
    sheet.getRange(row, 3)
      .setFormula(`=COUNTIF('MVR_Ticket_History'!G:G,"${status}")`)
      .setBackground(bgColor)
      .setNumberFormat('#,##0');
    
    sheet.getRange(row, 4)
      .setFormula(`=IFERROR(C${row}/SUM($C$${statusDataStart}:$C$${statusDataStart + 3}),0)`)
      .setBackground(bgColor)
      .setNumberFormat('0.0%');
    
    // Mini sparkline for trend
    sheet.getRange(row, 5)
      .setFormula(`=SPARKLINE({COUNTIFS('MVR_Ticket_History'!G:G,"${status}",'MVR_Ticket_History'!J:J,">="&TODAY()-90,'MVR_Ticket_History'!J:J,"<"&TODAY()-60),COUNTIFS('MVR_Ticket_History'!G:G,"${status}",'MVR_Ticket_History'!J:J,">="&TODAY()-60,'MVR_Ticket_History'!J:J,"<"&TODAY()-30),COUNTIFS('MVR_Ticket_History'!G:G,"${status}",'MVR_Ticket_History'!J:J,">="&TODAY()-30)},{"charttype","line";"color","#2563eb"})`)
      .setBackground(bgColor);
    
    row++;
  });
  
  row += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION: TOP PARTNERS (LIVE TABLE)
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(row, 2, 1, 4).merge()
    .setValue('🏢 TOP 10 PARTNERS BY VOLUME')
    .setBackground(COLORS.CAPACITY_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(row, 32);
  row++;
  
  // Populate top partners by volume using deduped Turn ID aggregation (in-script)
  createPartnerVolumeView(sheet, row);
  // createPartnerVolumeView will advance the row by writing 11 rows (header + 10) + spacing
  row += 12;

  // All partners full table (row-level)
  const usedAllPartners = createAllPartnersView(sheet, row, 2);
  row += (usedAllPartners || 20) + 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION: OUTCOME DISTRIBUTION (RIGHT SIDE)
  // ═══════════════════════════════════════════════════════════════════════════
  // Write time series for last 6 weeks (daily + weekly) before right-side widgets
  if (agg && agg.dailySeries && agg.weeklySeries) {
    writeTimeSeriesTable(sheet, row, 2, agg.dailySeries, agg.weeklySeries);
  }
  row += 44;

  const rightCol = 7;
  let rightRow = 4;
  
  sheet.getRange(rightRow, rightCol, 1, 3).merge()
    .setValue('🎯 OUTCOME DISTRIBUTION')
    .setBackground(COLORS.TIER_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(rightRow, 32);
  rightRow++;
  
  // Outcome distribution (computed from deduped Turn IDs)
  createOutcomeDistributionView(sheet, rightRow, rightCol, _rawData, _hmap);
  rightRow += 14;
  
  // Add State view to right side (by Turn ID deduped counts)
  createStateView(sheet, rightRow, rightCol);
  rightRow += 8;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION: KEY INSIGHTS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(rightRow, rightCol, 1, 3).merge()
    .setValue('💡 KEY INSIGHTS')
    .setBackground(COLORS.AMBER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  sheet.setRowHeight(rightRow, 32);
  rightRow++;
  
  // Dynamic insights using formulas
  const insights = [
    {
      condition: '=IF(COUNTIFS(\'MVR_Ticket_History\'!G:G,"Open",\'MVR_Ticket_History\'!M:M,">48")>5,"⚠️ "&COUNTIFS(\'MVR_Ticket_History\'!G:G,"Open",\'MVR_Ticket_History\'!M:M,">48")&" tickets open >48 hours - needs attention","✅ Open ticket backlog within SLA")',
      color: COLORS.WHITE
    },
    {
      condition: '=IF(COUNTIF(\'MVR_Ticket_History\'!AI:AI,"*Suspend*")>0,"🔴 "&COUNTIF(\'MVR_Ticket_History\'!AI:AI,"*Suspend*")&" suspensions detected - risk mitigation active","ℹ️ No suspensions in current dataset")',
      color: COLORS.WHITE
    },
    {
      condition: '=IF(COUNTIFS(\'MVR_Ticket_History\'!AA:AA,"UNKNOWN")>COUNTIFS(\'MVR_Ticket_History\'!AA:AA,"<>")*0.1,"⚠️ "&TEXT(COUNTIFS(\'MVR_Ticket_History\'!AA:AA,"UNKNOWN")/COUNTIFS(\'MVR_Ticket_History\'!AA:AA,"<>"),"0%")&" tickets have unknown vendor - check data quality","✅ Vendor classification above 90%")',
      color: COLORS.WHITE
    }
  ];
  
  // Render key insights using row-level data (avoid fragile sheet formulas)
  renderKeyInsights(sheet, rightRow, rightCol, _rawData, _hmap);
  rightRow += 3;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // FOOTER
  // ═══════════════════════════════════════════════════════════════════════════
  row = Math.max(row, rightRow) + 2;
  sheet.getRange(row, 2, 1, 8).merge()
    .setValue('📊 This dashboard updates automatically when data refreshes. All values reference live data from MVR_Ticket_History.')
    .setFontColor(COLORS.SLATE)
    .setFontSize(9)
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Freeze header
  sheet.setFrozenRows(2);
  
  // Create charts
  createStatusPieChart(sheet, statusDataStart);
  createMonthlyTrendChart(sheet, row);
  // Weekly financial chart is available as an auxiliary function but omitted from
  // the default CEO dashboard to keep the view operationally-focused. To add
  // the chart manually, call `createWeeklyFinancialChart(sheet, row)`.
  
  ui.alert('✅ CEO Dashboard Created',
    'Executive dashboard includes:\n' +
    '• 8 KPI cards with live calculations\n' +
    '• Status distribution with trends\n' +
    '• Top 10 partners by volume\n' +
    '• Outcome distribution\n' +
    '• Dynamic insights\n\n' +
    'All data is live and updates automatically.',
    ui.ButtonSet.OK);
}

/**
 * Create an executive-style KPI card with value, label, and optional trend
 */
function createExecutiveKPI(sheet, row, col, config) {
  // Card background
  sheet.getRange(row, col, 4, 1)
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, false, false, config.isAlert ? COLORS.RED : COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  
  // Colored top bar
  sheet.getRange(row, col)
    .setBackground(config.color)
    .setValue('');
  sheet.setRowHeight(row, 6);
  
  // Value (supports either formula string `config.value` or numeric `config.valueNumber`)
  const valueCell = sheet.getRange(row + 1, col);
  if (typeof config.valueNumber !== 'undefined') {
    valueCell.setValue(config.valueNumber);
  } else if (config.value) {
    valueCell.setFormula(config.value);
  } else {
    valueCell.setValue('');
  }
  valueCell
    .setFontSize(config.isPercent ? 24 : 26)
    .setFontWeight('bold')
    .setFontColor(config.isAlert ? COLORS.RED : COLORS.NAVY)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat(config.format);
  sheet.setRowHeight(row + 1, 45);
  
  // Label
  sheet.getRange(row + 2, col)
    .setValue(config.label)
    .setFontSize(10)
    .setFontColor(COLORS.SLATE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('top');
  
  // Trend or sparkline
  if (config.trend) {
    sheet.getRange(row + 3, col)
      .setFormula(`=IF(${config.trend}>0,"▲ "&TEXT(${config.trend},"0.0%"),"▼ "&TEXT(ABS(${config.trend}),"0.0%"))`)
      .setFontSize(9)
      .setHorizontalAlignment('center');
    // Color based on trend direction
    sheet.getRange(row + 3, col).setFontColor(COLORS.EMERALD); // Will be overwritten by conditional format
  }
  sheet.setRowHeight(row + 3, 20);
}

/**
 * Create status distribution pie chart
 */
function createStatusPieChart(sheet, startRow) {
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(`B${startRow}:B${startRow + 3}`))  // Labels
    .addRange(sheet.getRange(`C${startRow}:C${startRow + 3}`))  // Values
    .setPosition(startRow, 6, 20, 0)
    .setOption('title', 'Status Distribution')
    .setOption('titleTextStyle', { color: COLORS.NAVY, fontSize: 12, bold: true })
    .setOption('pieHole', 0.35)
    .setOption('colors', ['#fbbf24', '#3b82f6', '#10b981', '#6b7280'])
    .setOption('legend', { position: 'right', textStyle: { fontSize: 9 } })
    .setOption('width', 280)
    .setOption('height', 180);
  
  sheet.insertChart(chartBuilder.build());
}

// =========================
// In-script aggregation helpers (use Finance standards)
// =========================

/**
 * Safely parse a date-like value. Uses Date constructor then falls back to parseIsoDate.
 */
function safeParseDate(v) {
  if (!v && v !== 0) return null;
  if (v instanceof Date) return v;
  var d = new Date(v);
  if (!isNaN(d.getTime())) return d;
  try {
    return parseIsoDate(String(v));
  } catch (e) {
    return null;
  }
}

/**
 * Load history data and header map
 */
function loadHistoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  if (!sheet) throw new Error('MVR_Ticket_History sheet not found');
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { headers: [], data: [] };
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const headerMap = {};
  headers.forEach(function(h, i) { headerMap[h] = i; });
  return { sheet: sheet, headers: headers, headerMap: headerMap, data: data };
}

/**
 * Dedupe rows by Turn ID (canonical unit). Returns an array of rows (arrays).
 * Policy: prefer rows with `Outcome Override` present, otherwise choose the row
 * with the latest `Last Updated` (falling back to `Date Created`).
 */
function dedupeByTurnId(data, headerMap) {
  if (!data || data.length === 0) return [];
  const turnIdx = headerMap['Turn ID'];
  const ticketIdx = headerMap['Ticket ID'];
  const lastUpdatedIdx = headerMap['Last Updated'];
  const dateCreatedIdx = headerMap['Date Created'];
  const overrideIdx = headerMap['Outcome Override'];

  const map = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rawTurn = (typeof turnIdx !== 'undefined') ? (row[turnIdx] || '') : '';
    let key = String(rawTurn).toString().trim();
    if (!key) {
      // fallback grouping by Ticket ID when Turn ID missing
      key = 'TICKET:' + String(row[ticketIdx] || '').toString();
    }

    const existing = map[key];
    if (!existing) {
      map[key] = row;
      continue;
    }

    const existingOverride = (existing[overrideIdx] || '').toString().trim();
    const rowOverride = (row[overrideIdx] || '').toString().trim();

    // If current row has an override and existing doesn't, prefer current
    if (rowOverride && !existingOverride) { map[key] = row; continue; }

    // If both have overrides, choose the later Last Updated
    const existingDate = safeParseDate(existing[lastUpdatedIdx] || existing[dateCreatedIdx]);
    const rowDate = safeParseDate(row[lastUpdatedIdx] || row[dateCreatedIdx]);
    if (!existingDate && rowDate) { map[key] = row; continue; }
    if (existingDate && rowDate && rowDate.getTime() > existingDate.getTime()) { map[key] = row; continue; }
    // otherwise keep existing
  }

  const out = [];
  for (const k in map) out.push(map[k]);
  return out;
}

/**
 * Determine if a history row is billable using Tag_Outcome_Mappings (Finance standard)
 */
function rowIsBillable(row, headerMap, outcomeAnalysis) {
  // outcome column
  const outcomeIdx = headerMap['MVR Outcome'];
  const tagsIdx = headerMap['Tags'];
  const outcome = (row[outcomeIdx] || '').toString();
  const tags = (row[tagsIdx] || '').toString();

  for (let i = 0; i < outcomeAnalysis.length; i++) {
    const cfg = outcomeAnalysis[i];
    if (!cfg.isBillable) continue;
    if (cfg.outcome && cfg.outcome === outcome) return true;
    // For 'Clear' style outcomes, match tag patterns
    if (cfg.patterns && cfg.patterns.length > 0) {
      for (let j = 0; j < cfg.patterns.length; j++) {
        const p = cfg.patterns[j];
        try {
          if (tags && tags.toUpperCase().indexOf(p.toUpperCase()) !== -1) return true;
        } catch (e) {}
      }
    }
  }
  return false;
}

/**
 * Get vendor cost fallback: tries named range PARAM_COST_PER_CHECK_<VENDOR>, then FINANCE_COSTS_SHEET if present, otherwise defaults
 */
function getVendorCost(vendor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'PARAM_COST_PER_CHECK_' + String(vendor).toUpperCase();
  try {
    const r = ss.getRangeByName(name);
    if (r) return Number(r.getValue()) || 0;
  } catch (e) {}

  // Fallback defaults
  const defaults = { CERTN: 4.5, INFORMDATA: 3.75, PENNDOT: 0 };
  return defaults[vendor] || 0;
}

/**
 * Get Monday-start date for the week containing `d` (local timezone)
 */
function getWeekStart(d) {
  const dt = new Date(d);
  const day = (dt.getDay() + 6) % 7; // Monday=0..Sunday=6
  dt.setDate(dt.getDate() - day);
  dt.setHours(0,0,0,0);
  return dt;
}

/**
 * Compute CEO metrics using Finance aggregation standards
 * options: { days: 30, months: 6 }
 */
function computeCEOMetrics(options) {
  options = options || {};
  const days = options.days || 30;
  const months = options.months || 6;
  const history = loadHistoryData();
  const data = history.data;
  const h = history.headerMap;
  const today = new Date();
  const cutoff = new Date(today.getFullYear(), today.getMonth(), today.getDate() - days + 1);

  // Use raw row-level data for Operations dashboard (no dedupe)
  const dedupedData = data;

  const outcomeAnalysis = loadOutcomeAnalysisFromMappings();

  let totalTickets30 = 0;
  let billable30 = 0;
  let vendorCounts = {};
  let resolutionSum = 0;
  let resolutionCount = 0;
  let detectionMatches = 0;
  let outcomeNonEmpty = 0;
  let partnersSet = {};

  // Monthly buckets for financial trend (YYYY-MM -> { tickets, billable, revenue, cost })
  const monthly = {};
  // Daily (YYYY-MM-DD) and Weekly (week start YYYY-MM-DD) buckets
  const daily = {};
  const weekly = {};

  for (let i = 0; i < dedupedData.length; i++) {
    const row = dedupedData[i];
    const lastUpdated = safeParseDate(row[h['Last Updated']]);
    const createdDate = safeParseDate(row[h['Date Created']]);
    const useDate = lastUpdated || createdDate;
    if (!useDate) continue;

    // Monthly key
    const ymKey = formatYearMonthKey(useDate);
    if (!monthly[ymKey]) monthly[ymKey] = { tickets: 0, billable: 0, revenue: 0, cost: 0 };
    monthly[ymKey].tickets++;

    // Daily and weekly keys
    const tz = Session.getScriptTimeZone();
    const dayKey = Utilities.formatDate(useDate, tz, 'yyyy-MM-dd');
    if (!daily[dayKey]) daily[dayKey] = { tickets: 0, billable: 0, revenue: 0, cost: 0 };
    daily[dayKey].tickets++;
    const weekStart = getWeekStart(useDate);
    const weekKey = Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd');
    if (!weekly[weekKey]) weekly[weekKey] = { tickets: 0, billable: 0, revenue: 0, cost: 0 };
    weekly[weekKey].tickets++;

    // Overall 30-day window check (use Last Updated per Finance standards)
    const lastForWindow = lastUpdated || createdDate;
    if (lastForWindow && lastForWindow >= cutoff) {
      totalTickets30++;
      // partner
      const partner = (row[h['Partner Name']] || '').toString();
      if (partner) partnersSet[partner] = true;
      // resolution
      const res = Number(row[h['Resolution Time (Hours)']] || 0);
      if (res > 0) { resolutionSum += res; resolutionCount++; }
      // detection
      const outcome = (row[h['MVR Outcome']] || '').toString();
      if (outcome && outcome.toUpperCase().indexOf('SUSPEND') !== -1) detectionMatches++;
      if (outcome && outcome !== '') outcomeNonEmpty++;
      // vendor counts
      const vendor = (row[h['Vendor Group']] || 'UNKNOWN').toString();
      vendorCounts[vendor] = (vendorCounts[vendor] || 0) + 1;

      // billable check (for 30-day window counts)
      const isBill = rowIsBillable(row, h, outcomeAnalysis);
      if (isBill) {
        billable30++;
        const costPer = getVendorCost(vendor);
        monthly[ymKey].billable++;
        monthly[ymKey].revenue += costPer;
        monthly[ymKey].cost += costPer;
        // also accumulate daily/weekly billable/revenue for this row
        daily[dayKey].billable++;
        daily[dayKey].revenue += costPer;
        daily[dayKey].cost += costPer;
        weekly[weekKey].billable++;
        weekly[weekKey].revenue += costPer;
        weekly[weekKey].cost += costPer;
      }
    }
    // also accumulate monthly cost by vendor for full-month view
    // use vendor cost for any billable row regardless of window
    const vendorRow = (row[h['Vendor Group']] || 'UNKNOWN').toString();
    if (rowIsBillable(row, h, outcomeAnalysis)) {
      const cp = getVendorCost(vendorRow);
      monthly[ymKey].revenue += cp;
      monthly[ymKey].cost += cp;
      // For rows outside 30-day window we still account billable to daily/weekly totals
      const tz2 = Session.getScriptTimeZone();
      const dk = Utilities.formatDate(useDate, tz2, 'yyyy-MM-dd');
      const wk = Utilities.formatDate(getWeekStart(useDate), tz2, 'yyyy-MM-dd');
      if (!daily[dk]) daily[dk] = { tickets: 0, billable: 0, revenue: 0, cost: 0 };
      daily[dk].billable += 1;
      daily[dk].revenue += cp;
      daily[dk].cost += cp;
      if (!weekly[wk]) weekly[wk] = { tickets: 0, billable: 0, revenue: 0, cost: 0 };
      weekly[wk].billable += 1;
      weekly[wk].revenue += cp;
      weekly[wk].cost += cp;
    }
  }

  const totalPartners = Object.keys(partnersSet).length;
  const avgResolutionHours = resolutionCount ? (resolutionSum / resolutionCount) : 0;
  const detectionRate = outcomeNonEmpty ? (detectionMatches / outcomeNonEmpty) : 0;
  const slaCompliance = (function() {
    // calculate SLA compliance across deduped rows where Resolution Time is present
    let good = 0, total = 0;
    for (let i = 0; i < dedupedData.length; i++) {
      const row = dedupedData[i];
      const res = Number(row[h['Resolution Time (Hours)']] || 0);
      if (res > 0) { total++; if (res <= 24) good++; }
    }
    return total ? (good / total) : 0;
  })();

  // Build monthly arrays sorted ascending (old->new)
  const monthKeys = Object.keys(monthly).sort();
  const monthlySeries = monthKeys.map(k => ({ key: k, data: monthly[k] }));

  // Build daily series for last 42 days (6 weeks)
  const dailySeries = [];
  const daysBack = 42;
  for (let d = daysBack - 1; d >= 0; d--) {
    const dt = new Date();
    dt.setHours(0,0,0,0);
    dt.setDate(dt.getDate() - d);
    const k = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const val = daily[k] || { tickets: 0, billable: 0, revenue: 0, cost: 0 };
    dailySeries.push({ key: k, data: val });
  }

  // Build weekly series for last 6 weeks (Monday-start)
  const weeklySeries = [];
  const currentWeekStart = getWeekStart(new Date());
  for (let w = -5; w <= 0; w++) {
    const wkStart = new Date(currentWeekStart);
    wkStart.setDate(wkStart.getDate() + (w * 7));
    const k = Utilities.formatDate(wkStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const val = weekly[k] || { tickets: 0, billable: 0, revenue: 0, cost: 0 };
    weeklySeries.push({ key: k, data: val });
  }

  return {
    totalTickets30: totalTickets30,
    billable30: billable30,
    vendorCounts: vendorCounts,
    avgResolutionHours: avgResolutionHours,
    detectionRate: detectionRate,
    slaCompliance: slaCompliance,
    totalPartners: totalPartners,
    monthlySeries: monthlySeries,
    dailySeries: dailySeries,
    weeklySeries: weeklySeries,
    ticketsTrend: 0 // placeholder for backward-compat trend (can compute if needed)
  };
}

/**
 * Create a drill-down sheet filtered by a provided predicate on rows
 * predicate(row, headerMap) -> boolean
 */
function createDrillDownSheet(namePrefix, predicate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = loadHistoryData();
  const headers = h.headers;
  const rows = [headers];
  for (let i = 0; i < h.data.length; i++) {
    const r = h.data[i];
    if (predicate(r, h.headerMap)) rows.push(r);
  }
  const sheetName = namePrefix + ' - Drill';
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);
  const out = ss.insertSheet(sheetName);
  if (rows.length > 0) out.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  out.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('Drill-down sheet created: ' + sheetName);
  return out;
}


/**
 * Create monthly trend line chart
 */
function createMonthlyTrendChart(sheet, insertRow) {
  // First create the data for the chart in a hidden area
  const dataRow = 50;
  
  // Monthly labels and values
  for (let i = 5; i >= 0; i--) {
    sheet.getRange(dataRow, 2 + (5 - i))
      .setFormula(`=TEXT(EDATE(TODAY(),-${i}),"MMM")`);
    sheet.getRange(dataRow + 1, 2 + (5 - i))
      .setFormula(`=COUNTIFS('MVR_Ticket_History'!J:J,">="&EOMONTH(TODAY(),-${i + 1})+1,'MVR_Ticket_History'!J:J,"<="&EOMONTH(TODAY(),-${i}))`);
  }
  
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sheet.getRange(`B${dataRow}:G${dataRow}`))      // Labels
    .addRange(sheet.getRange(`B${dataRow + 1}:G${dataRow + 1}`))  // Values
    .setPosition(insertRow - 10, 2, 0, 0)
    .setOption('title', '6-Month Volume Trend')
    .setOption('titleTextStyle', { color: COLORS.NAVY, fontSize: 12, bold: true })
    .setOption('colors', [COLORS.ROYAL_BLUE])
    .setOption('legend', { position: 'none' })
    .setOption('curveType', 'function')
    .setOption('lineWidth', 3)
    .setOption('pointSize', 6)
    .setOption('width', 400)
    .setOption('height', 200)
    .setOption('hAxis', { textStyle: { fontSize: 10 } })
    .setOption('vAxis', { textStyle: { fontSize: 10 }, minValue: 0 });
  
  sheet.insertChart(chartBuilder.build());
}

/**
 * Create weekly financial series (last month and current month) and chart.
 * Weeks are 1..5 by day ranges: 1-7,8-14,15-21,22-28,29-31
 */
function createWeeklyFinancialChart(sheet, insertRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth(); // 0-based
  const lastMonthDate = new Date(currentYear, currentMonth - 1, 1);
  const lastMonth = lastMonthDate.getMonth();
  const lastYear = lastMonthDate.getFullYear();

  // Compute series
  const series = computeWeeklyFinancialSeries();

  // Write series to a hidden area on the CEO sheet (row 60) so charts can reference it
  const dataRow = 60;
  // Header row
  const headers = ['Week', formatMonthLabel(lastYear, lastMonth), formatMonthLabel(currentYear, currentMonth)];
  sheet.getRange(dataRow, 2, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(COLORS.MEDIUM_GRAY).setFontColor(COLORS.WHITE);

  // Weeks 1..5
  const rows = [];
  for (let w = 1; w <= 5; w++) {
    const lastVal = series.last[w - 1] || 0;
    const curVal = series.current[w - 1] || 0;
    rows.push(['W' + w, lastVal, curVal]);
  }

  sheet.getRange(dataRow + 1, 2, rows.length, 3).setValues(rows).setNumberFormat('$#,##0.00');

  // Build chart
  const chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(`B${dataRow + 1}:B${dataRow + 5}`)) // week labels
    .addRange(sheet.getRange(`C${dataRow + 1}:C${dataRow + 5}`)) // last month
    .addRange(sheet.getRange(`D${dataRow + 1}:D${dataRow + 5}`)) // current month
    .setPosition(insertRow - 8, 6, 0, 0)
    .setOption('title', 'Weekly Financials — Last vs Current Month')
    .setOption('width', 520)
    .setOption('height', 220)
    .setOption('legend', { position: 'right' })
    .setOption('hAxis', { title: 'Week' })
    .setOption('vAxis', { title: 'Revenue', format: '$#,##0' });

  sheet.insertChart(chartBuilder.build());
}

/**
 * Format month label like: "Mar 2025"
 */
function formatMonthLabel(year, monthZeroBased) {
  const d = new Date(year, monthZeroBased, 1);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy');
}

/**
 * Compute weekly financial series for last and current month.
 * Returns { last: [num,num,..], current: [num,num,..] } where index 0 => week1
 */
function computeWeeklyFinancialSeries() {
  const history = loadHistoryData();
  const data = history.data;
  const h = history.headerMap;
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  const lastMonthDate = new Date(currentYear, currentMonth - 1, 1);
  const lastYear = lastMonthDate.getFullYear();
  const lastMonth = lastMonthDate.getMonth();

  // Initialize five-week buckets
  const lastBuckets = [0, 0, 0, 0, 0];
  const currentBuckets = [0, 0, 0, 0, 0];

  const outcomeAnalysis = loadOutcomeAnalysisFromMappings();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const d = safeParseDate(row[h['Last Updated']] || row[h['Date Created']]);
    if (!d) continue;
    const y = d.getFullYear();
    const m = d.getMonth();
    const day = d.getDate();
    const weekIdx = Math.min(4, Math.ceil(day / 7) - 1 + 0); // 0..4

    // Determine if billable
    const isBill = rowIsBillable(row, h, outcomeAnalysis);
    if (!isBill) continue;

    const vendor = (row[h['Vendor Group']] || 'UNKNOWN').toString();
    const cost = getVendorCost(vendor) || 0;

    if (y === lastYear && m === lastMonth) {
      lastBuckets[weekIdx] = (lastBuckets[weekIdx] || 0) + cost;
    } else if (y === currentYear && m === currentMonth) {
      currentBuckets[weekIdx] = (currentBuckets[weekIdx] || 0) + cost;
    }
  }

  return { last: lastBuckets, current: currentBuckets };
}

/**
 * Create a State summary view (State, Tickets, Billable, % Billable) using row-level data (no dedupe).
 * Writes to the provided start row and column area (right-side usage).
 */
function createStateView(sheet, startRow, startCol) {
  const history = loadHistoryData();
  const h = history.headerMap;
  const data = history.data;
  const outcomeAnalysis = loadOutcomeAnalysisFromMappings();

  const stateIdx = h['DL State'];
  const vendorIdx = h['Vendor Group'];

  const map = {}; // state -> { tickets, billable }
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const state = (row[stateIdx] || 'UNKNOWN').toString() || 'UNKNOWN';
    if (!map[state]) map[state] = { tickets: 0, billable: 0 };
    map[state].tickets++;
    if (rowIsBillable(row, h, outcomeAnalysis)) map[state].billable++;
  }

  // Convert to rows sorted by tickets desc
  const rows = Object.keys(map).map(s => [s, map[s].tickets, map[s].billable, map[s].tickets ? (map[s].billable / map[s].tickets) : 0]);
  rows.sort((a, b) => b[1] - a[1]);

  // Header
  sheet.getRange(startRow, startCol, 1, 3).merge();
  sheet.getRange(startRow, startCol).setValue('📍 STATE SUMMARY').setBackground(COLORS.TIER_HEADER).setFontColor(COLORS.WHITE).setFontWeight('bold').setFontSize(12);
  sheet.setRowHeight(startRow, 28);

  // Table headers
  const headerRow = startRow + 1;
  sheet.getRange(headerRow, startCol, 1, 4).setValues([['State','Tickets','Billable','% Billable']])
    .setBackground(COLORS.MEDIUM_GRAY).setFontColor(COLORS.WHITE).setFontWeight('bold');

  const startDataRow = headerRow + 1;
  const writeRows = rows.slice(0, 12); // limit to top 12 states for compact view
  if (writeRows.length > 0) {
    sheet.getRange(startDataRow, startCol, writeRows.length, 4).setValues(writeRows);
    sheet.getRange(startDataRow, startCol + 3, writeRows.length, 1).setNumberFormat('0.0%');
  }

  // Style box
  sheet.getRange(startRow, startCol, (writeRows.length + 3), 4)
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Create Top Partners by Volume view using row-level data (no dedupe).
 * Writes partner, tickets, avg resolution (hrs) for top 10 partners starting at given row.
 */
function createPartnerVolumeView(sheet, startRow) {
  const history = loadHistoryData();
  const h = history.headerMap;
  const data = history.data;

  const partnerIdx = h['Partner Name'];
  const resIdx = h['Resolution Time (Hours)'];

  const map = {}; // partner -> { tickets, resSum }
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = (row[partnerIdx] || '').toString().trim() || 'UNKNOWN';
    if (!map[p]) map[p] = { tickets: 0, resSum: 0, resCount: 0 };
    map[p].tickets++;
    const r = Number(row[resIdx] || 0);
    if (r > 0) { map[p].resSum += r; map[p].resCount++; }
  }

  const rows = Object.keys(map).map(p => [p, map[p].tickets, map[p].resCount ? (map[p].resSum / map[p].resCount) : 0]);
  rows.sort((a, b) => b[1] - a[1]);
  const top = rows.slice(0, 10);

  // Header already written by caller; write table header and data
  const headerRow = startRow;
  sheet.getRange(headerRow, 2, 1, 3).setValues([['Partner','Tickets','Avg Res (hrs)']])
    .setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');

  const dataRow = headerRow + 1;
  if (top.length > 0) sheet.getRange(dataRow, 2, top.length, 3).setValues(top).setNumberFormat(['@','#,##0','0.0']);

  // Style box
  sheet.getRange(headerRow, 2, (top.length + 1), 3)
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Create Outcome distribution view from row-level data (no dedupe).
 * Writes Outcome and Count starting at provided cell (startRow, startCol).
 */
function createOutcomeDistributionView(sheet, startRow, startCol, data, headerMap) {
  const outcomeIdx = headerMap['MVR Outcome'];
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const o = (data[i][outcomeIdx] || '').toString().trim() || 'Unknown';
    map[o] = (map[o] || 0) + 1;
  }

  const rows = Object.keys(map).map(k => [k, map[k]]).sort((a,b) => b[1] - a[1]);

  // Header
  sheet.getRange(startRow, startCol, 1, 2).setValues([['Outcome','Count']])
    .setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');

  const dataStart = startRow + 1;
  const writeRows = rows.slice(0, 12);
  if (writeRows.length > 0) sheet.getRange(dataStart, startCol, writeRows.length, 2).setValues(writeRows);

  // Style
  sheet.getRange(startRow, startCol, (writeRows.length + 1), 2)
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Compute and render key insights as plain text (avoid fragile formulas).
 */
function renderKeyInsights(sheet, startRow, startCol, data, headerMap) {
  const statusIdx = headerMap['Status'];
  const ageIdx = headerMap['Age (Hours)'] || headerMap['Resolution Time (Hours)'];
  const outcomeIdx = headerMap['MVR Outcome'];
  const vendorIdx = headerMap['Vendor Group'];

  // Insight 1: open tickets >48 hours
  let openOver48 = 0;
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    const status = (r[statusIdx]||'').toString().trim();
    const age = Number(r[ageIdx]||0);
    if (status === 'Open' && age > 48) openOver48++;
  }

  // Insight 2: suspensions count
  let suspCount = 0;
  for (let i = 0; i < data.length; i++) {
    const o = (data[i][outcomeIdx]||'').toString();
    if (o.toLowerCase().indexOf('suspend') !== -1) suspCount++;
  }

  // Insight 3: vendor unknown rate
  let unknownVendor = 0; let totalV = 0;
  for (let i = 0; i < data.length; i++) {
    const v = (data[i][vendorIdx]||'').toString();
    if (v === '' || v.toUpperCase() === 'UNKNOWN') unknownVendor++;
    if (v !== '') totalV++;
  }
  const unknownRate = totalV ? (unknownVendor / totalV) : 0;

  // Write insights
  const insights = [
    openOver48 > 0 ? `⚠️ ${openOver48} tickets open >48 hours - needs attention` : '✅ Open ticket backlog within SLA',
    suspCount > 0 ? `🔴 ${suspCount} suspensions detected - risk mitigation active` : 'ℹ️ No suspensions in current dataset',
    unknownRate > 0.1 ? `⚠️ ${(unknownRate*100).toFixed(0)}% tickets have unknown vendor - check data quality` : '✅ Vendor classification above 90%'
  ];

  for (let i = 0; i < insights.length; i++) {
    sheet.getRange(startRow + i, startCol, 1, 3).merge().setValue(insights[i]).setWrap(true).setFontSize(10).setBackground(COLORS.WHITE);
    sheet.setRowHeight(startRow + i, 35);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create an "All Partners" table (row-level): Partner, Tickets, Billable, Avg Res (hrs), Total Cost
 */
function createAllPartnersView(sheet, startRow, startCol) {
  const history = loadHistoryData();
  const h = history.headerMap;
  const data = history.data;
  const outcomeAnalysis = loadOutcomeAnalysisFromMappings();

  const partnerIdx = h['Partner Name'];
  const resIdx = h['Resolution Time (Hours)'];
  const vendorIdx = h['Vendor Group'];

  const map = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const p = (row[partnerIdx] || '').toString().trim() || 'UNKNOWN';
    if (!map[p]) map[p] = { tickets: 0, billable: 0, resSum: 0, resCount: 0, cost: 0 };
    map[p].tickets++;
    const r = Number(row[resIdx] || 0);
    if (r > 0) { map[p].resSum += r; map[p].resCount++; }
    const isBill = rowIsBillable(row, h, outcomeAnalysis);
    if (isBill) {
      map[p].billable++;
      const vendor = (row[vendorIdx] || 'UNKNOWN').toString();
      const cp = getVendorCost(vendor) || 0;
      map[p].cost += cp;
    }
  }

  const rows = Object.keys(map).map(p => [p, map[p].tickets, map[p].billable, map[p].resCount ? (map[p].resSum / map[p].resCount) : 0, map[p].cost]);
  rows.sort((a,b) => b[1] - a[1]);

  // Header
  sheet.getRange(startRow, startCol, 1, 5).setValues([['Partner','Tickets','Billable','Avg Res (hrs)','Total Cost']])
    .setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');

  if (rows.length > 0) sheet.getRange(startRow + 1, startCol, rows.length, 5).setValues(rows).setNumberFormat(['@','#,##0','#,##0','0.0','$#,##0.00']);

  sheet.getRange(startRow, startCol, (rows.length + 1), 5)
    .setBackground(COLORS.WHITE)
    .setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  return rows.length + 1; // header + data rows
}

/**
 * Write contiguous time series tables for daily (42 days) and weekly (6 weeks) series.
 */
function writeTimeSeriesTable(sheet, startRow, startCol, dailySeries, weeklySeries) {
  // Daily table
  sheet.getRange(startRow, startCol, 1, 4).setValues([['Date','Tickets','Billable','Revenue']])
    .setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');
  const dailyRows = dailySeries.map(s => [s.key, s.data.tickets || 0, s.data.billable || 0, s.data.revenue || 0]);
  if (dailyRows.length > 0) sheet.getRange(startRow + 1, startCol, dailyRows.length, 4).setValues(dailyRows).setNumberFormat(['@','#,##0','#,##0','$#,##0.00']);

  // Weekly table placed to the right
  const wkCol = startCol + 6;
  sheet.getRange(startRow, wkCol, 1, 4).setValues([['WeekStart','Tickets','Billable','Revenue']])
    .setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');
  const weeklyRows = weeklySeries.map(s => [s.key, s.data.tickets || 0, s.data.billable || 0, s.data.revenue || 0]);
  if (weeklyRows.length > 0) sheet.getRange(startRow + 1, wkCol, weeklyRows.length, 4).setValues(weeklyRows).setNumberFormat(['@','#,##0','#,##0','$#,##0.00']);

  // Style boxes
  sheet.getRange(startRow, startCol, (dailyRows.length + 1), 4).setBackground(COLORS.WHITE).setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(startRow, wkCol, (weeklyRows.length + 1), 4).setBackground(COLORS.WHITE).setBorder(true, true, true, true, true, true, COLORS.BORDER_GRAY, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Menu entry point for creating CEO Dashboard
 */
function runCreateCEODashboard() {
  createCEODashboard();
}
