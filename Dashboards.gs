/**
 * DASHBOARD SYSTEM
 * MVR TICKET TRACKER - Multi-Dimensional Analytics Views
 * 
 * Purpose: Generate dashboard views from Partner, Vendor, State, and Geographic perspectives
 * Plus monthly reconciliation for Finance billing
 */

// ═══════════════════════════════════════════════════════════════════════════════
// US REGIONS (for Geographic Dashboard)
// ═══════════════════════════════════════════════════════════════════════════════

const US_REGIONS = {
  NORTHEAST: ["CT", "ME", "MA", "NH", "RI", "VT", "NJ", "NY", "PA"],
  MIDWEST: ["IL", "IN", "MI", "OH", "WI", "IA", "KS", "MN", "MO", "NE", "ND", "SD"],
  SOUTH: ["DE", "FL", "GA", "MD", "NC", "SC", "VA", "DC", "WV", "AL", "KY", "MS", "TN", "AR", "LA", "OK", "TX"],
  WEST: ["AZ", "CO", "ID", "MT", "NV", "NM", "UT", "WY", "AK", "CA", "HI", "OR", "WA"]
};

// ═══════════════════════════════════════════════════════════════════════════════
// DASHBOARD DATA LOADER
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Load ticket history data with column mapping
 * @return {Object} { data: Array, colMap: Object, headers: Array }
 */
function loadDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { data: [], colMap: {}, headers: [] };
  }
  
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const data = allData.slice(1);
  
  // Build column map
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);
  
  return { data, colMap, headers };
}

/**
 * Get or create dashboard sheet
 */
function getOrCreateDashboardSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  
  if (headers && headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#4a86e8")
      .setFontColor("white");
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════════════════
// PARTNER DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Partner Dashboard - View from customer perspective
 */
function generatePartnerDashboard() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('📊 Generating Partner Dashboard...');
    
    const { data, colMap } = loadDashboardData();
    if (data.length === 0) {
      ui.alert('No Data', 'No ticket history found. Run a refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    // Aggregate by partner
    const partnerStats = {};
    
    data.forEach(row => {
      const partner = row[colMap['Partner Name']] || 'Unknown';
      const type = row[colMap['Request Type (SC/RC)']] || '';
      const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
      const state = row[colMap['DL State']] || '';
      const outcome = row[colMap['MVR Outcome']] || '';
      const resHours = parseFloat(row[colMap['Resolution Time (Hours)']]) || 0;
      // Parse ISO date strings for correct calculations
      const createdDate = row[colMap['Date Created']] ? new Date(row[colMap['Date Created']]) : null;
      const lastUpdatedDate = row[colMap['Last Updated']] ? new Date(row[colMap['Last Updated']]) : null;
      // For all action tracking, use createdDate as start and lastUpdatedDate as end
      // For billing, use only lastUpdatedDate
      const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
      
      if (!partnerStats[partner]) {
        partnerStats[partner] = {
          total: 0, sc: 0, rc: 0,
          certn: 0, penndot: 0, informdata: 0, other: 0,
          states: {}, billable: 0, totalHours: 0
        };
      }
      
      const p = partnerStats[partner];
      p.total++;
      if (type === 'SC') p.sc++;
      if (type === 'RC') p.rc++;
      if (vendor === 'CERTN') p.certn++;
      else if (vendor === 'PENNDOT') p.penndot++;
      else if (vendor === 'INFORMDATA') p.informdata++;
      else p.other++;
      if (state) p.states[state] = (p.states[state] || 0) + 1;
      if (isBillable) p.billable++;
      if (resHours > 0) p.totalHours += resHours;
    });
    
    // Build output
    const headers = [
      'Partner', 'Total', 'SC', 'RC', 
      'CERTN', 'PENNDOT', 'INFORMDATA', 'Other Vendor',
      'Top States', 'Billable %', 'Avg Resolution (hrs)'
    ];
    
    const rows = Object.entries(partnerStats)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([partner, stats]) => {
        // Top 3 states
        const topStates = Object.entries(stats.states)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 3)
          .map(([s, c]) => `${s}(${c})`)
          .join(', ');
        
        const billablePct = stats.total > 0 ? 
          ((stats.billable / stats.total) * 100).toFixed(1) + '%' : '0%';
        const avgHours = stats.billable > 0 ? 
          (stats.totalHours / stats.billable).toFixed(1) : '-';
        
        return [
          partner, stats.total, stats.sc, stats.rc,
          stats.certn, stats.penndot, stats.informdata, stats.other,
          topStates, billablePct, avgHours
        ];
      });
    
    // Write to sheet
    const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_PARTNERS, headers);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Add summary row
    const totalRow = rows.reduce((acc, row) => {
      acc[1] += row[1]; acc[2] += row[2]; acc[3] += row[3];
      acc[4] += row[4]; acc[5] += row[5]; acc[6] += row[6]; acc[7] += row[7];
      return acc;
    }, ['TOTAL', 0, 0, 0, 0, 0, 0, 0, '-', '-', '-']);
    
    sheet.getRange(rows.length + 2, 1, 1, headers.length)
      .setValues([totalRow])
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✅ Partner Dashboard: ${rows.length} partners`);
    ui.alert('✅ Partner Dashboard Generated', 
      `${rows.length} partners analyzed.\n\nView: ${SHEET_NAMES.DASHBOARD_PARTNERS}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Partner Dashboard error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// VENDOR DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Vendor Dashboard - View from vendor/cost perspective
 */
function generateVendorDashboard() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('📊 Generating Vendor Dashboard...');
    
    const { data, colMap } = loadDashboardData();
    if (data.length === 0) {
      ui.alert('No Data', 'No ticket history found. Run a refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    // Aggregate by vendor
    const vendorStats = {};
    const months = {};
    
    data.forEach(row => {
      const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
      const partner = row[colMap['Partner Name']] || 'Unknown';
      const state = row[colMap['DL State']] || '';
      const outcome = row[colMap['MVR Outcome']] || '';
      const created = row[colMap['Date Created']];
      // Parse ISO date string for correct calculations
      const createdDate = created ? new Date(created) : null;
      const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
      // Monthly tracking
      if (createdDate) {
        const monthKey = `${createdDate.getFullYear()}-${String(createdDate.getMonth() + 1).padStart(2, '0')}`;
        v.monthly[monthKey] = (v.monthly[monthKey] || 0) + 1;
        months[monthKey] = true;
      }
      
      // Determine FREE/PAID
      const isFree = state && STATE_CONFIG[state] && STATE_CONFIG[state].is_free;
      
      if (!vendorStats[vendor]) {
        vendorStats[vendor] = {
          total: 0, free: 0, paid: 0,
          partners: new Set(), states: new Set(),
          billable: 0, monthly: {}
        };
      }
      
      const v = vendorStats[vendor];
      v.total++;
      if (isFree) v.free++; else v.paid++;
      v.partners.add(partner);
      if (state) v.states.add(state);
      if (isBillable) v.billable++;
      
      // Monthly tracking
      if (created) {
        const d = new Date(created);
        const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
        v.monthly[monthKey] = (v.monthly[monthKey] || 0) + 1;
        months[monthKey] = true;
      }
    });
    
    // Get sorted months for columns
    const sortedMonths = Object.keys(months).sort().slice(-6); // Last 6 months
    
    // Build output
    const headers = [
      'Vendor', 'Total', 'FREE', 'PAID', 
      'Partners', 'States Covered', 'Billable',
      ...sortedMonths.map(m => m.substring(5)) // Just MM
    ];
    
    const rows = Object.entries(vendorStats)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([vendor, stats]) => {
        const monthlyVals = sortedMonths.map(m => stats.monthly[m] || 0);
        
        return [
          vendor, stats.total, stats.free, stats.paid,
          stats.partners.size, stats.states.size, stats.billable,
          ...monthlyVals
        ];
      });
    
    // Write to sheet
    const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_VENDORS, headers);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✅ Vendor Dashboard: ${rows.length} vendors`);
    ui.alert('✅ Vendor Dashboard Generated', 
      `${rows.length} vendors analyzed.\n\nView: ${SHEET_NAMES.DASHBOARD_VENDORS}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Vendor Dashboard error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// STATE DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate State Dashboard - View from jurisdiction perspective
 */
function generateStateDashboard() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('📊 Generating State Dashboard...');
    
    const { data, colMap } = loadDashboardData();
    if (data.length === 0) {
      ui.alert('No Data', 'No ticket history found. Run a refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    // Aggregate by state
    const stateStats = {};
    
    data.forEach(row => {
      const state = row[colMap['DL State']] || 'Unknown';
      const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
      const partner = row[colMap['Partner Name']] || 'Unknown';
      const outcome = row[colMap['MVR Outcome']] || '';
      
      if (!stateStats[state]) {
        stateStats[state] = {
          total: 0, vendors: {}, partners: new Set(),
          cleared: 0, suspended: 0, expired: 0, other: 0, billable: 0
        };
      }
      
      const s = stateStats[state];
      s.total++;
      s.vendors[vendor] = (s.vendors[vendor] || 0) + 1;
      s.partners.add(partner);
      
      // Outcome breakdown
      const outLower = outcome.toLowerCase();
      if (outLower.includes('clear')) s.cleared++;
      else if (outLower.includes('suspend')) s.suspended++;
      else if (outLower.includes('expir')) s.expired++;
      else s.other++;
      
      const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
      if (isBillable) s.billable++;
    });
    
    // Build output
    const headers = [
      'State', 'Primary Vendor', 'FREE/PAID', 'Total',
      'Partners', 'Billable', 'CLEARED', 'SUSPENDED', 'EXPIRED', 'OTHER'
    ];
    
    const rows = Object.entries(stateStats)
      .filter(([state]) => state !== 'Unknown')
      .sort((a, b) => b[1].total - a[1].total)
      .map(([state, stats]) => {
        // Primary vendor (most volume)
        const primaryVendor = Object.entries(stats.vendors)
          .sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A';
        
        // FREE/PAID from config
        const freePaid = STATE_CONFIG[state]?.is_free ? 'FREE' : 'PAID';
        
        return [
          state, primaryVendor, freePaid, stats.total,
          stats.partners.size, stats.billable,
          stats.cleared, stats.suspended, stats.expired, stats.other
        ];
      });
    
    // Write to sheet
    const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_STATES, headers);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Color FREE states
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][2] === 'FREE') {
        sheet.getRange(i + 2, 3).setBackground('#d9ead3'); // Light green
      }
    }
    
    sheet.autoResizeColumns(1, headers.length);
    
    Logger.log(`✅ State Dashboard: ${rows.length} states`);
    ui.alert('✅ State Dashboard Generated', 
      `${rows.length} states analyzed.\n\nView: ${SHEET_NAMES.DASHBOARD_STATES}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ State Dashboard error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// GEOGRAPHIC DASHBOARD
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Geographic Dashboard - Regional analysis
 */
function generateGeographicDashboard() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    Logger.log('📊 Generating Geographic Dashboard...');
    
    const { data, colMap } = loadDashboardData();
    if (data.length === 0) {
      ui.alert('No Data', 'No ticket history found. Run a refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    // Map state to region
    const stateToRegion = {};
    Object.entries(US_REGIONS).forEach(([region, states]) => {
      states.forEach(s => stateToRegion[s] = region);
    });
    
    // Aggregate by region
    const regionStats = { NORTHEAST: {}, MIDWEST: {}, SOUTH: {}, WEST: {}, OTHER: {} };
    Object.keys(regionStats).forEach(r => {
      regionStats[r] = { total: 0, free: 0, paid: 0, partners: new Set(), vendors: {}, states: {} };
    });
    
    data.forEach(row => {
      const state = row[colMap['DL State']] || '';
      const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
      const partner = row[colMap['Partner Name']] || 'Unknown';
      const isFree = state && STATE_CONFIG[state] && STATE_CONFIG[state].is_free;
      
      const region = stateToRegion[state] || 'OTHER';
      const r = regionStats[region];
      
      r.total++;
      if (isFree) r.free++; else r.paid++;
      r.partners.add(partner);
      r.vendors[vendor] = (r.vendors[vendor] || 0) + 1;
      if (state) r.states[state] = (r.states[state] || 0) + 1;
    });
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD_GEOGRAPHIC);
    if (sheet) sheet.clearContents();
    else sheet = ss.insertSheet(SHEET_NAMES.DASHBOARD_GEOGRAPHIC);
    
    let currentRow = 1;
    
    // Title
    sheet.getRange(currentRow, 1).setValue('GEOGRAPHIC ANALYSIS DASHBOARD')
      .setFontSize(14).setFontWeight('bold');
    sheet.getRange(currentRow, 2).setValue(`Generated: ${new Date().toLocaleString()}`);
    currentRow += 2;
    
    // Section 1: Regional Summary
    sheet.getRange(currentRow, 1).setValue('REGIONAL SUMMARY')
      .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    currentRow++;
    
    const regionHeaders = ['Region', 'Total', 'FREE', 'PAID', 'Partners', 'Primary Vendor', 'Top States'];
    sheet.getRange(currentRow, 1, 1, regionHeaders.length).setValues([regionHeaders])
      .setFontWeight('bold').setBackground('#e8f0fe');
    currentRow++;
    
    ['NORTHEAST', 'MIDWEST', 'SOUTH', 'WEST', 'OTHER'].forEach(region => {
      const r = regionStats[region];
      const primaryVendor = Object.entries(r.vendors).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A';
      const topStates = Object.entries(r.states).sort((a, b) => b[1] - a[1])
        .slice(0, 3).map(([s, c]) => `${s}(${c})`).join(', ') || 'N/A';
      
      sheet.getRange(currentRow, 1, 1, regionHeaders.length).setValues([[
        region, r.total, r.free, r.paid, r.partners.size, primaryVendor, topStates
      ]]);
      currentRow++;
    });
    
    currentRow += 2;
    
    // Section 2: State Heat Map Data
    sheet.getRange(currentRow, 1).setValue('STATE VOLUME (Top 25)')
      .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    currentRow++;
    
    const stateVolumes = {};
    data.forEach(row => {
      const state = row[colMap['DL State']] || '';
      if (state) stateVolumes[state] = (stateVolumes[state] || 0) + 1;
    });
    
    const stateHeaders = ['State', 'Region', 'Volume', 'FREE/PAID'];
    sheet.getRange(currentRow, 1, 1, stateHeaders.length).setValues([stateHeaders])
      .setFontWeight('bold').setBackground('#e8f0fe');
    currentRow++;
    
    Object.entries(stateVolumes)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 25)
      .forEach(([state, volume]) => {
        const region = stateToRegion[state] || 'OTHER';
        const freePaid = STATE_CONFIG[state]?.is_free ? 'FREE' : 'PAID';
        sheet.getRange(currentRow, 1, 1, stateHeaders.length).setValues([[
          state, region, volume, freePaid
        ]]);
        if (freePaid === 'FREE') {
          sheet.getRange(currentRow, 4).setBackground('#d9ead3');
        }
        currentRow++;
      });
    
    currentRow += 2;
    
    // Section 3: Vendor by Region
    sheet.getRange(currentRow, 1).setValue('VENDOR DISTRIBUTION BY REGION')
      .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    currentRow++;
    
    const vendorRegionHeaders = ['Region', 'CERTN', 'PENNDOT', 'INFORMDATA', 'OTHER'];
    sheet.getRange(currentRow, 1, 1, vendorRegionHeaders.length).setValues([vendorRegionHeaders])
      .setFontWeight('bold').setBackground('#e8f0fe');
    currentRow++;
    
    ['NORTHEAST', 'MIDWEST', 'SOUTH', 'WEST'].forEach(region => {
      const v = regionStats[region].vendors;
      sheet.getRange(currentRow, 1, 1, vendorRegionHeaders.length).setValues([[
        region, v.CERTN || 0, v.PENNDOT || 0, v.INFORMDATA || 0,
        (v.LEGACY || 0) + (v.UNKNOWN || 0) + (v.INFORM || 0)
      ]]);
      currentRow++;
    });
    
    sheet.autoResizeColumns(1, 10);
    
    Logger.log('✅ Geographic Dashboard generated');
    ui.alert('✅ Geographic Dashboard Generated', 
      `Regional analysis complete.\n\nView: ${SHEET_NAMES.DASHBOARD_GEOGRAPHIC}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Geographic Dashboard error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// MONTHLY RECONCILIATION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate Monthly Reconciliation for Finance - Transaction-level detail
 * @param {number} year - Year to generate (defaults to current year)
 */
function generateMonthlyReconciliation(year = null) {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Default to current year
    if (!year) {
      year = new Date().getFullYear();
    }
    
    Logger.log(`📊 Generating Monthly Reconciliation for ${year}...`);
    
    const { data, colMap } = loadDashboardData();
    if (data.length === 0) {
      ui.alert('No Data', 'No ticket history found. Run a refresh first.', ui.ButtonSet.OK);
      return;
    }
    
    // Filter to completed tickets in the target year
    const yearData = data.filter(row => {
      const resolved = row[colMap['Date Resolved']];
      const outcome = row[colMap['MVR Outcome']] || '';
      if (!resolved || !outcome) return false;
      const d = new Date(resolved);
      return d.getFullYear() === year && !['STILL_PROCESSING', 'PENDING'].includes(outcome);
    });
    
    // Group by month
    const monthlyData = {};
    const monthNames = ['January', 'February', 'March', 'April', 'May', 'June',
                        'July', 'August', 'September', 'October', 'November', 'December'];
    
    yearData.forEach(row => {
      const resolved = new Date(row[colMap['Date Resolved']]);
      const month = resolved.getMonth(); // 0-11
      
      if (!monthlyData[month]) monthlyData[month] = [];
      
      const state = row[colMap['DL State']] || '';
      const outcome = row[colMap['MVR Outcome']] || '';
      const isBillable = outcome && !['UNKNOWN', 'RECORD_NOT_FOUND', 'CANCELLED'].includes(outcome);
      
      monthlyData[month].push({
        date: resolved.toISOString().split('T')[0],
        partner: row[colMap['Partner Name']] || '',
        turnId: row[colMap['Turn ID']] || '',
        vendor: row[colMap['Vendor Group']] || '',
        state: state,
        outcome: outcome,
        billable: isBillable ? '✓' : ''
      });
    });
    
    // Create sheet
    const sheetName = SHEET_NAMES.RECONCILIATION_PREFIX + year;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) sheet.clearContents();
    else sheet = ss.insertSheet(sheetName);
    
    let currentRow = 1;
    const headers = ['Date Completed', 'Partner', 'Turn ID', 'Vendor', 'State', 'Outcome', 'Billable'];
    
    // Title
    sheet.getRange(currentRow, 1).setValue(`MONTHLY RECONCILIATION - ${year}`)
      .setFontSize(14).setFontWeight('bold');
    sheet.getRange(currentRow, 3).setValue(`Generated: ${new Date().toLocaleString()}`);
    currentRow += 2;
    
    // Process each month
    for (let m = 0; m < 12; m++) {
      const transactions = monthlyData[m] || [];
      if (transactions.length === 0 && m > new Date().getMonth() && year === new Date().getFullYear()) {
        continue; // Skip future months in current year
      }
      
      // Month header
      const billableCount = transactions.filter(t => t.billable === '✓').length;
      sheet.getRange(currentRow, 1, 1, 7).merge();
      sheet.getRange(currentRow, 1)
        .setValue(`${monthNames[m].toUpperCase()} ${year}   |   Total: ${transactions.length}   |   Billable: ${billableCount}`)
        .setFontWeight('bold')
        .setBackground('#4a86e8')
        .setFontColor('white');
      currentRow++;
      
      if (transactions.length > 0) {
        // Column headers
        sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers])
          .setFontWeight('bold').setBackground('#e8f0fe');
        currentRow++;
        
        // Sort by date, then partner
        transactions.sort((a, b) => {
          if (a.date !== b.date) return a.date.localeCompare(b.date);
          return a.partner.localeCompare(b.partner);
        });
        
        // Data rows
        const rows = transactions.map(t => [
          t.date, t.partner, t.turnId, t.vendor, t.state, t.outcome, t.billable
        ]);
        sheet.getRange(currentRow, 1, rows.length, headers.length).setValues(rows);
        currentRow += rows.length;
      } else {
        sheet.getRange(currentRow, 1).setValue('No completed transactions')
          .setFontStyle('italic');
        currentRow++;
      }
      
      currentRow++; // Blank row between months
    }
    
    // Year-end summary
    currentRow++;
    sheet.getRange(currentRow, 1, 1, 7).merge();
    sheet.getRange(currentRow, 1)
      .setValue(`YEAR SUMMARY ${year}`)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('white');
    currentRow++;
    
    const totalTransactions = yearData.length;
    const totalBillable = yearData.filter(row => {
      const outcome = row[colMap['MVR Outcome']] || '';
      return outcome && !['UNKNOWN', 'RECORD_NOT_FOUND', 'CANCELLED', 'STILL_PROCESSING', 'PENDING'].includes(outcome);
    }).length;
    
    // Summary by vendor
    const vendorSummary = {};
    yearData.forEach(row => {
      const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
      vendorSummary[vendor] = (vendorSummary[vendor] || 0) + 1;
    });
    
    sheet.getRange(currentRow, 1).setValue('Total Transactions:');
    sheet.getRange(currentRow, 2).setValue(totalTransactions).setFontWeight('bold');
    currentRow++;
    sheet.getRange(currentRow, 1).setValue('Total Billable:');
    sheet.getRange(currentRow, 2).setValue(totalBillable).setFontWeight('bold');
    currentRow++;
    currentRow++;
    
    sheet.getRange(currentRow, 1).setValue('By Vendor:').setFontWeight('bold');
    currentRow++;
    Object.entries(vendorSummary).sort((a, b) => b[1] - a[1]).forEach(([vendor, count]) => {
      sheet.getRange(currentRow, 1).setValue(`  ${vendor}:`);
      sheet.getRange(currentRow, 2).setValue(count);
      currentRow++;
    });
    
    sheet.autoResizeColumns(1, 7);
    sheet.setFrozenRows(0); // No frozen rows for this report
    
    Logger.log(`✅ Monthly Reconciliation ${year}: ${totalTransactions} transactions`);
    ui.alert('✅ Monthly Reconciliation Generated', 
      `Year: ${year}\nTransactions: ${totalTransactions}\nBillable: ${totalBillable}\n\nView: ${sheetName}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Monthly Reconciliation error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Prompt for year and generate reconciliation
 */
function runMonthlyReconciliation() {
  const currentYear = new Date().getFullYear();
  generateMonthlyReconciliation(currentYear);
}

// ═══════════════════════════════════════════════════════════════════════════════
// REFRESH ALL DASHBOARDS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Generate all dashboards at once
 */
function refreshAllDashboards() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '🔄 Refresh All Dashboards',
    'This will refresh/regenerate:\n\n' +
    '📈 LIVE DASHBOARDS (formula-based):\n' +
    '• Finance Dashboard\n' +
    '• CEO Dashboard\n' +
    '• Pivot Analysis\n\n' +
    '📊 STATIC DASHBOARDS:\n' +
    '• Partner Dashboard\n' +
    '• Vendor Dashboard\n' +
    '• State Dashboard\n' +
    '• Geographic Dashboard\n' +
    '• Monthly Reconciliation\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  try {
    Logger.log('🔄 Refreshing all dashboards...');
    
    // Step 1: Force recalculation of live dashboards
    Logger.log('📈 Refreshing live dashboards (formulas)...');
    SpreadsheetApp.flush();
    
    // Update Assumptions timestamp if it exists
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assumptions = ss.getSheetByName('Assumptions');
    if (assumptions) {
      const data = assumptions.getDataRange().getValues();
      for (let i = 0; i < Math.min(5, data.length); i++) {
        for (let j = 0; j < data[i].length; j++) {
          if (String(data[i][j]).includes('Last Updated')) {
            assumptions.getRange(i + 1, j + 1).setValue(`Last Updated: ${new Date().toLocaleString()}`);
            break;
          }
        }
      }
    }
    
    // Step 2: Regenerate static dashboards (suppress individual alerts)
    Logger.log('📊 Regenerating static dashboards...');
    generatePartnerDashboardSilent();
    generateVendorDashboardSilent();
    generateStateDashboardSilent();
    generateGeographicDashboardSilent();
    generateMonthlyReconciliationSilent();
    
    ui.alert('✅ All Dashboards Refreshed',
      '📈 Live dashboards (formulas recalculated):\n' +
      '• Finance Dashboard\n' +
      '• CEO Dashboard\n' +
      '• Pivot Analysis\n\n' +
      '📊 Static dashboards regenerated:\n' +
      `• ${SHEET_NAMES.DASHBOARD_PARTNERS}\n` +
      `• ${SHEET_NAMES.DASHBOARD_VENDORS}\n` +
      `• ${SHEET_NAMES.DASHBOARD_STATES}\n` +
      `• ${SHEET_NAMES.DASHBOARD_GEOGRAPHIC}\n` +
      `• ${SHEET_NAMES.RECONCILIATION_PREFIX}${new Date().getFullYear()}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Refresh all error: ${e.message}`);
    ui.alert('Error', e.message, ui.ButtonSet.OK);
  }
}

// Silent versions (no UI alerts) for batch refresh
function generatePartnerDashboardSilent() {
  const { data, colMap } = loadDashboardData();
  if (data.length === 0) return;
  
  const partnerStats = {};
  data.forEach(row => {
    const partner = row[colMap['Partner Name']] || 'Unknown';
    const type = row[colMap['Request Type (SC/RC)']] || '';
    const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
    const state = row[colMap['DL State']] || '';
    const outcome = row[colMap['MVR Outcome']] || '';
    const resHours = parseFloat(row[colMap['Resolution Time (Hours)']]) || 0;
    const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
    
    if (!partnerStats[partner]) {
      partnerStats[partner] = {
        total: 0, sc: 0, rc: 0, certn: 0, penndot: 0, informdata: 0, other: 0,
        states: {}, billable: 0, totalHours: 0
      };
    }
    
    const p = partnerStats[partner];
    p.total++;
    if (type === 'SC') p.sc++;
    if (type === 'RC') p.rc++;
    if (vendor === 'CERTN') p.certn++;
    else if (vendor === 'PENNDOT') p.penndot++;
    else if (vendor === 'INFORMDATA') p.informdata++;
    else p.other++;
    if (state) p.states[state] = (p.states[state] || 0) + 1;
    if (isBillable) p.billable++;
    if (resHours > 0) p.totalHours += resHours;
  });
  
  const headers = ['Partner', 'Total', 'SC', 'RC', 'CERTN', 'PENNDOT', 'INFORMDATA', 'Other Vendor', 'Top States', 'Billable %', 'Avg Resolution (hrs)'];
  const rows = Object.entries(partnerStats).sort((a, b) => b[1].total - a[1].total).map(([partner, stats]) => {
    const topStates = Object.entries(stats.states).sort((a, b) => b[1] - a[1]).slice(0, 3).map(([s, c]) => `${s}(${c})`).join(', ');
    const billablePct = stats.total > 0 ? ((stats.billable / stats.total) * 100).toFixed(1) + '%' : '0%';
    const avgHours = stats.billable > 0 ? (stats.totalHours / stats.billable).toFixed(1) : '-';
    return [partner, stats.total, stats.sc, stats.rc, stats.certn, stats.penndot, stats.informdata, stats.other, topStates, billablePct, avgHours];
  });
  
  const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_PARTNERS, headers);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

function generateVendorDashboardSilent() {
  const { data, colMap } = loadDashboardData();
  if (data.length === 0) return;
  
  const vendorStats = {};
  const months = {};
  
  data.forEach(row => {
    const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
    const partner = row[colMap['Partner Name']] || 'Unknown';
    const state = row[colMap['DL State']] || '';
    const outcome = row[colMap['MVR Outcome']] || '';
    const created = row[colMap['Date Created']];
    const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
    const isFree = state && STATE_CONFIG[state] && STATE_CONFIG[state].is_free;
    
    if (!vendorStats[vendor]) {
      vendorStats[vendor] = { total: 0, free: 0, paid: 0, partners: new Set(), states: new Set(), billable: 0, monthly: {} };
    }
    
    const v = vendorStats[vendor];
    v.total++;
    if (isFree) v.free++; else v.paid++;
    v.partners.add(partner);
    if (state) v.states.add(state);
    if (isBillable) v.billable++;
    
    if (created) {
      const d = new Date(created);
      const monthKey = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      v.monthly[monthKey] = (v.monthly[monthKey] || 0) + 1;
      months[monthKey] = true;
    }
  });
  
  const sortedMonths = Object.keys(months).sort().slice(-6);
  const headers = ['Vendor', 'Total', 'FREE', 'PAID', 'Partners', 'States Covered', 'Billable', ...sortedMonths.map(m => m.substring(5))];
  const rows = Object.entries(vendorStats).sort((a, b) => b[1].total - a[1].total).map(([vendor, stats]) => {
    const monthlyVals = sortedMonths.map(m => stats.monthly[m] || 0);
    return [vendor, stats.total, stats.free, stats.paid, stats.partners.size, stats.states.size, stats.billable, ...monthlyVals];
  });
  
  const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_VENDORS, headers);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

function generateStateDashboardSilent() {
  const { data, colMap } = loadDashboardData();
  if (data.length === 0) return;
  
  const stateStats = {};
  data.forEach(row => {
    const state = row[colMap['DL State']] || 'Unknown';
    const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
    const partner = row[colMap['Partner Name']] || 'Unknown';
    const outcome = row[colMap['MVR Outcome']] || '';
    
    if (!stateStats[state]) {
      stateStats[state] = { total: 0, vendors: {}, partners: new Set(), cleared: 0, suspended: 0, expired: 0, other: 0, billable: 0 };
    }
    
    const s = stateStats[state];
    s.total++;
    s.vendors[vendor] = (s.vendors[vendor] || 0) + 1;
    s.partners.add(partner);
    
    const outLower = outcome.toLowerCase();
    if (outLower.includes('clear')) s.cleared++;
    else if (outLower.includes('suspend')) s.suspended++;
    else if (outLower.includes('expir')) s.expired++;
    else s.other++;
    
    const isBillable = outcome && !['STILL_PROCESSING', 'PENDING', 'UNKNOWN'].includes(outcome);
    if (isBillable) s.billable++;
  });
  
  const headers = ['State', 'Primary Vendor', 'FREE/PAID', 'Total', 'Partners', 'Billable', 'CLEARED', 'SUSPENDED', 'EXPIRED', 'OTHER'];
  const rows = Object.entries(stateStats).filter(([state]) => state !== 'Unknown').sort((a, b) => b[1].total - a[1].total).map(([state, stats]) => {
    const primaryVendor = Object.entries(stats.vendors).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A';
    const freePaid = STATE_CONFIG[state]?.is_free ? 'FREE' : 'PAID';
    return [state, primaryVendor, freePaid, stats.total, stats.partners.size, stats.billable, stats.cleared, stats.suspended, stats.expired, stats.other];
  });
  
  const sheet = getOrCreateDashboardSheet(SHEET_NAMES.DASHBOARD_STATES, headers);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);
}

function generateGeographicDashboardSilent() {
  // Simplified - just call the main one without alerts
  const { data, colMap } = loadDashboardData();
  if (data.length === 0) return;
  
  const stateToRegion = {};
  Object.entries(US_REGIONS).forEach(([region, states]) => {
    states.forEach(s => stateToRegion[s] = region);
  });
  
  const regionStats = { NORTHEAST: {}, MIDWEST: {}, SOUTH: {}, WEST: {}, OTHER: {} };
  Object.keys(regionStats).forEach(r => {
    regionStats[r] = { total: 0, free: 0, paid: 0, partners: new Set(), vendors: {}, states: {} };
  });
  
  data.forEach(row => {
    const state = row[colMap['DL State']] || '';
    const vendor = row[colMap['Vendor Group']] || 'UNKNOWN';
    const partner = row[colMap['Partner Name']] || 'Unknown';
    const isFree = state && STATE_CONFIG[state] && STATE_CONFIG[state].is_free;
    const region = stateToRegion[state] || 'OTHER';
    const r = regionStats[region];
    r.total++;
    if (isFree) r.free++; else r.paid++;
    r.partners.add(partner);
    r.vendors[vendor] = (r.vendors[vendor] || 0) + 1;
    if (state) r.states[state] = (r.states[state] || 0) + 1;
  });
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.DASHBOARD_GEOGRAPHIC);
  if (sheet) sheet.clearContents();
  else sheet = ss.insertSheet(SHEET_NAMES.DASHBOARD_GEOGRAPHIC);
  
  let currentRow = 1;
  sheet.getRange(currentRow, 1).setValue('GEOGRAPHIC ANALYSIS DASHBOARD').setFontSize(14).setFontWeight('bold');
  currentRow += 2;
  
  const regionHeaders = ['Region', 'Total', 'FREE', 'PAID', 'Partners', 'Primary Vendor'];
  sheet.getRange(currentRow, 1, 1, regionHeaders.length).setValues([regionHeaders]).setFontWeight('bold').setBackground('#e8f0fe');
  currentRow++;
  
  ['NORTHEAST', 'MIDWEST', 'SOUTH', 'WEST', 'OTHER'].forEach(region => {
    const r = regionStats[region];
    const primaryVendor = Object.entries(r.vendors).sort((a, b) => b[1] - a[1])[0]?.[0] || 'N/A';
    sheet.getRange(currentRow, 1, 1, regionHeaders.length).setValues([[region, r.total, r.free, r.paid, r.partners.size, primaryVendor]]);
    currentRow++;
  });
  
  sheet.autoResizeColumns(1, 10);
}

function generateMonthlyReconciliationSilent() {
  const year = new Date().getFullYear();
  const { data, colMap } = loadDashboardData();
  if (data.length === 0) return;
  
  const yearData = data.filter(row => {
    const resolved = row[colMap['Date Resolved']];
    const outcome = row[colMap['MVR Outcome']] || '';
    if (!resolved || !outcome) return false;
    const d = new Date(resolved);
    return d.getFullYear() === year && !['STILL_PROCESSING', 'PENDING'].includes(outcome);
  });
  
  const monthlyData = {};
  const monthNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  
  yearData.forEach(row => {
    const resolved = new Date(row[colMap['Date Resolved']]);
    const month = resolved.getMonth();
    if (!monthlyData[month]) monthlyData[month] = [];
    const state = row[colMap['DL State']] || '';
    const outcome = row[colMap['MVR Outcome']] || '';
    const isBillable = outcome && !['UNKNOWN', 'RECORD_NOT_FOUND', 'CANCELLED'].includes(outcome);
    monthlyData[month].push({
      date: resolved.toISOString().split('T')[0],
      partner: row[colMap['Partner Name']] || '',
      turnId: row[colMap['Turn ID']] || '',
      vendor: row[colMap['Vendor Group']] || '',
      state: state,
      outcome: outcome,
      billable: isBillable ? '✓' : ''
    });
  });
  
  const sheetName = SHEET_NAMES.RECONCILIATION_PREFIX + year;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) sheet.clearContents();
  else sheet = ss.insertSheet(sheetName);
  
  let currentRow = 1;
  const headers = ['Date Completed', 'Partner', 'Turn ID', 'Vendor', 'State', 'Outcome', 'Billable'];
  
  sheet.getRange(currentRow, 1).setValue(`MONTHLY RECONCILIATION - ${year}`).setFontSize(14).setFontWeight('bold');
  currentRow += 2;
  
  for (let m = 0; m < 12; m++) {
    const transactions = monthlyData[m] || [];
    if (transactions.length === 0 && m > new Date().getMonth()) continue;
    
    const billableCount = transactions.filter(t => t.billable === '✓').length;
    sheet.getRange(currentRow, 1, 1, 7).merge();
    sheet.getRange(currentRow, 1).setValue(`${monthNames[m].toUpperCase()} ${year}   |   Total: ${transactions.length}   |   Billable: ${billableCount}`)
      .setFontWeight('bold').setBackground('#4a86e8').setFontColor('white');
    currentRow++;
    
    if (transactions.length > 0) {
      sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
      currentRow++;
      transactions.sort((a, b) => a.date !== b.date ? a.date.localeCompare(b.date) : a.partner.localeCompare(b.partner));
      const rows = transactions.map(t => [t.date, t.partner, t.turnId, t.vendor, t.state, t.outcome, t.billable]);
      sheet.getRange(currentRow, 1, rows.length, headers.length).setValues(rows);
      currentRow += rows.length;
    }
    currentRow++;
  }
  
  sheet.autoResizeColumns(1, 7);
}
