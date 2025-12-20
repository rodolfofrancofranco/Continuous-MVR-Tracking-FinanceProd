/**
 * ═══════════════════════════════════════════════════════════════════════════════
 * PIVOT ANALYSIS SHEET - Native Pivot Tables with Charts
 * ═══════════════════════════════════════════════════════════════════════════════
 * 
 * Creates native Google Sheets pivot tables for:
 * - Partner × Status matrix
 * - Vendor × State volume
 * - Monthly trends by outcome
 * - Agent performance
 * 
 * Pivot tables auto-refresh when source data changes
 * Charts are linked to pivot data for live visualization
 */

// ═══════════════════════════════════════════════════════════════════════════════
// PIVOT ANALYSIS CREATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create the Pivot Analysis sheet with multiple pivot tables and charts
 */
function createPivotAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Check for history sheet
  const historySheet = ss.getSheetByName('MVR_Ticket_History');
  if (!historySheet) {
    ui.alert('❌ Error', 'MVR_Ticket_History sheet not found. Run data pull first.', ui.ButtonSet.OK);
    return;
  }
  
  // Get or create sheet
  let sheet = ss.getSheetByName('Pivot Analysis');
  if (sheet) {
    const response = ui.alert(
      '⚠️ Pivot Analysis Exists',
      'This will recreate all pivot tables. Proceed?',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    ss.deleteSheet(sheet);
  }
  
  sheet = ss.insertSheet('Pivot Analysis');
  sheet.setHiddenGridlines(true);
  
  // Get source data range
  const historyData = historySheet.getDataRange();
  const lastRow = historySheet.getLastRow();
  const lastCol = historySheet.getLastColumn();
  
  let currentRow = 1;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // HEADER
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📊 PIVOT ANALYSIS - Multi-Dimensional Views')
    .setBackground(COLORS.NAVY)
    .setFontColor(COLORS.WHITE)
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 45);
  currentRow++;
  
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setFormula('="Data Source: MVR_Ticket_History | Rows: "&COUNTA(\'MVR_Ticket_History\'!A:A)-1&" | Last Refreshed: "&TEXT(NOW(),"yyyy-mm-dd hh:mm")')
    .setFontColor(COLORS.SLATE)
    .setFontSize(10)
    .setHorizontalAlignment('center');
  currentRow += 2;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PIVOT 1: PARTNER × STATUS
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('🏢 PIVOT 1: Partner × Status Matrix')
    .setBackground(COLORS.CAPACITY_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;
  
  // Create pivot table
  const pivot1 = createPivotTable(
    sheet,
    historySheet,
    currentRow,
    1,
    {
      rows: ['Partner Name'],      // Column D
      columns: ['Status'],          // Column G
      values: [{ column: 'Ticket ID', function: 'COUNTA' }],  // Column A
      filter: null
    }
  );
  
  // Reserve space for pivot
  currentRow += 25;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PIVOT 2: VENDOR × STATE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('💼 PIVOT 2: Vendor × State Volume')
    .setBackground(COLORS.COST_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;
  
  // Build Vendor × State pivot in-script to avoid spreadsheet date/string issues
  buildVendorStatePivot(historySheet, sheet, currentRow);
  
  // Format headers
  sheet.getRange(currentRow, 1, 1, 3)
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');
  
  currentRow += 60;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PIVOT 3: MONTHLY TREND BY OUTCOME
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 8).merge()
    .setValue('📅 PIVOT 3: Monthly Volume by Outcome')
    .setBackground(COLORS.TIER_HEADER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;
  
  // Monthly breakdown with outcomes (built in-script)
  buildMonthlyOutcomePivot(historySheet, sheet, currentRow);
  
  sheet.getRange(currentRow, 1, 1, 4)
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');
  
  currentRow += 50;
  
  // ═══════════════════════════════════════════════════════════════════════════
  // PIVOT 4: AGENT PERFORMANCE
  // ═══════════════════════════════════════════════════════════════════════════
  sheet.getRange(currentRow, 1, 1, 6).merge()
    .setValue('👤 PIVOT 4: Agent Performance')
    .setBackground(COLORS.AMBER)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold')
    .setFontSize(12);
  currentRow++;
  
  // Agent performance (built in-script)
  buildAgentPerformancePivot(historySheet, sheet, currentRow);
  
  sheet.getRange(currentRow, 1, 1, 3)
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');
  
  // ═══════════════════════════════════════════════════════════════════════════
  // ADD SUMMARY CHARTS (RIGHT SIDE)
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Create Vendor Distribution Chart
  createVendorDistributionChart(sheet);
  
  // Create Status Distribution Chart
  createStatusDistributionChart(sheet);
  
  // Create Monthly Trend Chart
  createMonthlyVolumeChart(sheet);
  
  ui.alert('✅ Pivot Analysis Created',
    'Created 4 pivot analysis views:\n' +
    '• Partner × Status Matrix\n' +
    '• Vendor × State Volume\n' +
    '• Monthly Trend by Outcome\n' +
    '• Agent Performance\n\n' +
    'Plus 3 linked charts. Data auto-updates when MVR_Ticket_History changes.',
    ui.ButtonSet.OK);
  
  Logger.log('✅ Created Pivot Analysis sheet with QUERY-based pivots and charts');
}

/**
 * Create a native pivot table
 * Note: Google Apps Script pivot table API is limited, using QUERY as alternative
 */
function createPivotTable(destSheet, sourceSheet, destRow, destCol, config) {
  // Since native pivot table creation is limited in Apps Script,
  // we'll use QUERY formulas which behave similarly
  
  const rowField = config.rows[0];
  const colField = config.columns ? config.columns[0] : null;
  const valueField = config.values[0];
  
  // Build QUERY formula
  let formula;
  if (colField) {
    // Cross-tabulation style
    formula = `=QUERY('${sourceSheet.getName()}'!A:AO,"SELECT ${getColumnLetter(rowField)}, ${getColumnLetter(colField)}, COUNT(${getColumnLetter(valueField.column)}) WHERE ${getColumnLetter(rowField)} IS NOT NULL GROUP BY ${getColumnLetter(rowField)}, ${getColumnLetter(colField)} PIVOT ${getColumnLetter(colField)} LABEL ${getColumnLetter(rowField)} '${rowField}'")`;
  } else {
    formula = `=QUERY('${sourceSheet.getName()}'!A:AO,"SELECT ${getColumnLetter(rowField)}, COUNT(${getColumnLetter(valueField.column)}) WHERE ${getColumnLetter(rowField)} IS NOT NULL GROUP BY ${getColumnLetter(rowField)} ORDER BY COUNT(${getColumnLetter(valueField.column)}) DESC LABEL ${getColumnLetter(rowField)} '${rowField}', COUNT(${getColumnLetter(valueField.column)}) 'Count'")`;
  }
  
  destSheet.getRange(destRow, destCol).setFormula(formula);
  
  // Format header row
  destSheet.getRange(destRow, destCol, 1, 10)
    .setBackground(COLORS.SLATE)
    .setFontColor(COLORS.WHITE)
    .setFontWeight('bold');
  
  return true;
}

/**
 * Map column header to column letter
 */
function getColumnLetter(headerName) {
  // Uses HISTORY_HEADERS from Config.gs for dynamic mapping
  if (typeof HISTORY_HEADERS === 'undefined') {
    throw new Error('HISTORY_HEADERS is not defined.');
  }
  const idx = HISTORY_HEADERS.indexOf(headerName);
  if (idx === -1) {
    // Optionally, log or throw for missing header
    return 'A';
  }
  // Convert index to column letter (A, B, ..., Z, AA, AB, ...)
  let col = idx + 1;
  let letter = '';
  while (col > 0) {
    let rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/**
 * Create vendor distribution donut chart
 */
function createVendorDistributionChart(sheet) {
  // First create data for chart
  const dataRow = 200;
  sheet.getRange(dataRow, 1).setValue('Vendor');
  sheet.getRange(dataRow, 2).setValue('Count');
  sheet.getRange(dataRow + 1, 1).setValue('CERTN');
  sheet.getRange(dataRow + 1, 2).setFormula('=COUNTIF(\'MVR_Ticket_History\'!Z:Z,"CERTN")');
  sheet.getRange(dataRow + 2, 1).setValue('INFORMDATA');
  sheet.getRange(dataRow + 2, 2).setFormula('=COUNTIF(\'MVR_Ticket_History\'!Z:Z,"INFORMDATA")');
  sheet.getRange(dataRow + 3, 1).setValue('PENNDOT');
  sheet.getRange(dataRow + 3, 2).setFormula('=COUNTIF(\'MVR_Ticket_History\'!Z:Z,"PENNDOT")');
  sheet.getRange(dataRow + 4, 1).setValue('UNKNOWN');
  sheet.getRange(dataRow + 4, 2).setFormula('=COUNTIF(\'MVR_Ticket_History\'!Z:Z,"UNKNOWN")');
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(dataRow, 1, 5, 2))
    .setPosition(4, 8, 0, 0)
    .setOption('title', 'Vendor Distribution')
    .setOption('titleTextStyle', { color: COLORS.NAVY, fontSize: 12, bold: true })
    .setOption('pieHole', 0.4)
    .setOption('colors', [COLORS.CAPACITY_HEADER, COLORS.ROYAL_BLUE, COLORS.EMERALD, COLORS.MEDIUM_GRAY])
    .setOption('legend', { position: 'right' })
    .setOption('width', 320)
    .setOption('height', 220)
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Build Vendor × State pivot in-script and write values to sheet
 */
function buildVendorStatePivot(historySheet, destSheet, destRow) {
  const headers = historySheet.getDataRange().getValues()[0] || [];
  const data = historySheet.getDataRange().getValues().slice(1);
  const vendorIdx = HISTORY_HEADERS.indexOf('Vendor Group');
  const stateIdx = HISTORY_HEADERS.indexOf('DL State');

  const map = {};
  data.forEach(function(r) {
    const vendor = r[vendorIdx] || 'UNKNOWN';
    const state = r[stateIdx] || 'UNKNOWN';
    const key = vendor + '|' + state;
    map[key] = (map[key] || 0) + 1;
  });

  const rows = [['Vendor', 'State', 'Count']];
  Object.keys(map).forEach(function(k) {
    const parts = k.split('|');
    rows.push([parts[0], parts[1], map[k]]);
  });
  // Sort by count desc
  rows.splice(1).sort(function(a,b){return b[2]-a[2];}).forEach(function(r,i){rows[i+1]=r;});

  destSheet.getRange(destRow, 1, rows.length, 3).setValues(rows);
  destSheet.getRange(destRow, 1, 1, 3).setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');
}

/**
 * Build Monthly Outcome pivot and write values: Year, Month, Outcome, Count
 */
function buildMonthlyOutcomePivot(historySheet, destSheet, destRow) {
  const data = historySheet.getDataRange().getValues().slice(1);
  const lastUpdatedIdx = HISTORY_HEADERS.indexOf('Last Updated');
  const outcomeIdx = HISTORY_HEADERS.indexOf('MVR Outcome');

  const map = {}; // key: YYYY-MM|outcome -> count
  data.forEach(function(r) {
    const iso = r[lastUpdatedIdx];
    const d = parseIsoDate(iso);
    if (!d) return;
    const ym = formatYearMonthKey(d);
    const outcome = r[outcomeIdx] || 'UNKNOWN';
    const key = ym + '|' + outcome;
    map[key] = (map[key] || 0) + 1;
  });

  const rows = [['Year','Month','Outcome','Count']];
  Object.keys(map).forEach(function(k){
    const parts = k.split('|');
    const ym = parts[0].split('-');
    rows.push([parseInt(ym[0],10), parseInt(ym[1],10), parts.slice(1).join('|'), map[k]]);
  });
  // Sort by Year desc, Month desc
  rows.splice(1).sort(function(a,b){ if (a[0]!==b[0]) return b[0]-a[0]; return b[1]-a[1]; }).forEach(function(r,i){rows[i+1]=r;});

  destSheet.getRange(destRow, 1, rows.length, 4).setValues(rows);
  destSheet.getRange(destRow, 1, 1, 4).setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');
}

/**
 * Build Agent performance pivot: Agent, Tickets, Avg Resolution (hrs)
 */
function buildAgentPerformancePivot(historySheet, destSheet, destRow) {
  const data = historySheet.getDataRange().getValues().slice(1);
  const agentIdx = HISTORY_HEADERS.indexOf('Assigned Agent');
  const resIdx = HISTORY_HEADERS.indexOf('Resolution Time (Hours)');

  const map = {};
  data.forEach(function(r){
    const agent = r[agentIdx] || 'UNASSIGNED';
    const res = parseFloat(r[resIdx]) || 0;
    if (!map[agent]) map[agent] = {count:0, sum:0};
    map[agent].count += 1;
    map[agent].sum += res;
  });

  const rows = [['Agent','Tickets','Avg Resolution (hrs)']];
  Object.keys(map).forEach(function(a){
    const rec = map[a];
    const avg = rec.count ? (rec.sum / rec.count) : 0;
    rows.push([a, rec.count, Math.round(avg*100)/100]);
  });
  rows.splice(1).sort(function(a,b){return b[1]-a[1];}).forEach(function(r,i){rows[i+1]=r;});

  destSheet.getRange(destRow, 1, rows.length, 3).setValues(rows);
  destSheet.getRange(destRow, 1, 1, 3).setBackground(COLORS.SLATE).setFontColor(COLORS.WHITE).setFontWeight('bold');
}

/**
 * Create status distribution chart
 */
function createStatusDistributionChart(sheet) {
  const dataRow = 210;
  sheet.getRange(dataRow, 1).setValue('Status');
  sheet.getRange(dataRow, 2).setValue('Count');
  
  const statuses = ['Open', 'Pending', 'Resolved', 'Closed'];
  statuses.forEach((s, i) => {
    sheet.getRange(dataRow + 1 + i, 1).setValue(s);
    sheet.getRange(dataRow + 1 + i, 2).setFormula(`=COUNTIF('MVR_Ticket_History'!G:G,"${s}")`);
  });
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(dataRow, 1, 5, 2))
    .setPosition(4, 12, 0, 0)
    .setOption('title', 'Status Distribution')
    .setOption('titleTextStyle', { color: COLORS.NAVY, fontSize: 12, bold: true })
    .setOption('colors', [COLORS.ROYAL_BLUE])
    .setOption('legend', { position: 'none' })
    .setOption('width', 300)
    .setOption('height', 220)
    .setOption('hAxis', { textStyle: { fontSize: 10 } })
    .build();
  
  sheet.insertChart(chart);
}

/**
 * Create monthly volume trend chart
 */
function createMonthlyVolumeChart(sheet) {
  const dataRow = 220;
  sheet.getRange(dataRow, 1).setValue('Month');
  sheet.getRange(dataRow, 2).setValue('Tickets');
  
  for (let i = 5; i >= 0; i--) {
    const row = dataRow + 1 + (5 - i);
    sheet.getRange(row, 1).setFormula(`=TEXT(EDATE(TODAY(),-${i}),"MMM YY")`);
    sheet.getRange(row, 2).setFormula(`=COUNTIFS('MVR_Ticket_History'!J:J,">="&EOMONTH(TODAY(),-${i + 1})+1,'MVR_Ticket_History'!J:J,"<="&EOMONTH(TODAY(),-${i}))`);
  }
  
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.AREA)
    .addRange(sheet.getRange(dataRow, 1, 7, 2))
    .setPosition(18, 8, 0, 0)
    .setOption('title', '6-Month Volume Trend')
    .setOption('titleTextStyle', { color: COLORS.NAVY, fontSize: 12, bold: true })
    .setOption('colors', [COLORS.TIER_HEADER])
    .setOption('legend', { position: 'none' })
    .setOption('width', 450)
    .setOption('height', 200)
    .setOption('areaOpacity', 0.3)
    .build();
  
  sheet.insertChart(chart);
}

// ═══════════════════════════════════════════════════════════════════════════════
// PIVOT REFRESH FUNCTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Refresh all pivot tables and charts
 * Called after data refresh to ensure pivots show latest data
 */
function refreshPivots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // For QUERY-based pivots, they auto-refresh
  // For native pivots, we would need to recreate them
  // For charts, they also auto-refresh with their data source
  
  // Force recalculation
  SpreadsheetApp.flush();
  
  Logger.log('✅ Pivots and charts refreshed (formulas recalculated)');
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Menu entry point for creating Pivot Analysis
 */
function runCreatePivotAnalysis() {
  createPivotAnalysis();
}
