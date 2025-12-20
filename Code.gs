/**
 * MVR TICKET TRACKER - MAIN ENTRY POINT
 * Code.gs - Custom Menu, Orchestration, and Trigger Setup
 * 
 * Purpose: Main interface for MVR Ticket Tracker system
 * 
 * UNIFIED PIPELINE:
 * API → MVR_Raw_Tickets → MVR_Ticket_History → Reports
 * 
 * Hourly: Raw (2hr) + Process to History
 * Daily 9 AM: Raw (30 days) + Process to History + Reports
 */

// ═══════════════════════════════════════════════════════════════════════════════
// CUSTOM MENU - USER-FRIENDLY STRUCTURE
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Create custom menu when spreadsheet opens
 * Simplified structure for clarity
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 MVR Tracker')
    // PRIMARY ACTIONS - Most used
    .addItem('▶️ RUN COMPLETE REFRESH', 'runCompleteRefresh')
    .addSeparator()
    
    // REFRESH DATA - Pull from Freshdesk
    .addSubMenu(ui.createMenu('🔄 Refresh Data')
      .addItem('⚡ Quick Refresh (2 hours) - Fast update for recent changes', 'runQuickRefresh')
      .addItem('📦 Full Refresh (30 days) - Complete data pull', 'runFullRefresh')
      .addSeparator()
      .addItem('🔁 Reprocess All - Re-enrich from Raw to History', 'runReprocessAll')
      .addItem('🔄 Re-Enrich Data - Re-extract partner/state from notes', 'runReEnrichment')
      .addItem('🏷️ Refresh Vendor Groups - Fix vendor classification', 'runRefreshVendorGroups')
      .addSeparator()
      .addItem('✅ Validate Data Columns - Check report readiness', 'runValidateDataColumns')
      .addItem('🔍 Debug Unknown Vendors - Diagnose extraction issues', 'debugUnknownVendors')
      .addSeparator()
      .addItem('🆕 INITIAL SETUP - First-time pull (clears all data)', 'runInitialSetup'))
    
    // REPORTS - Generate outputs
    .addSubMenu(ui.createMenu('📈 Reports')
      .addItem('📊 Generate All Reports', 'runAllReports')
      .addSeparator()
      .addItem('⚙️ Setup Assumptions Sheet - Configure parameters', 'runCreateAssumptions')
      .addItem('📋 Assumptions Log - Central reference', 'runAssumptionsLog')
      .addSeparator()
      .addItem('🔄 OPS Lifecycle Report - Ticket flow & blockers', 'runOpsLifecycleReport')
      .addItem('💰 Finance Reconciliation - Vendor billing & Turn IDs', 'runFinanceReconciliationReport')
      .addItem('📈 CEO Effectiveness - Detection & outcomes', 'runCEOEffectivenessReport')
      .addSeparator()
      .addItem('📅 Annual Projection - Capacity planning', 'runAnnualProjection')
      .addItem('🗓️ Operative Plan - Wave calendar', 'runOperativePlan')
      .addItem('📈 Growth Projection - 5-year outlook', 'runGrowthProjection')
      .addSeparator()
      .addItem('🏢 Partner Summary - By company', 'runPartnerSummary'))
    
    // TAG MANAGEMENT - Discovery and mappings
    .addSubMenu(ui.createMenu('🏷️ Tag Management')
      .addItem('🔄 SYNC MAPPINGS: Load/Reset All Standard Patterns', 'syncTagMappings')
      .addSeparator()
      .addItem('🔍 View Discovered Tags - All tags with stats', 'viewDiscoveredTags')
      .addItem('⚠️ View Unmapped Tags - Need your decision', 'viewUnmappedTags')
      .addSeparator()
      .addItem('📋 View Current Mappings - All classification rules', 'viewTagMappings')
      .addItem('➕ Map Single Tag - Add one mapping', 'mapNewTag')
      .addItem('🧪 Test Classification - Preview tag outcome', 'testTagMatch')
      .addSeparator()
      .addItem('📊 Analyze Combinations - Pattern discovery (Advanced)', 'runTagCombinationAnalysis')
      .addSeparator()
      .addItem('🔄 Reapply All Rules - Reclassify history with current mappings', 'reapplyTagMappingsToHistory')
      .addSeparator()
      .addItem('🧪 Run QA Tests - Validate classification system', 'runTagClassificationQA')
      .addSeparator()
      .addItem('♻️ Reset to Defaults - Clear all and start over', 'resetTagMappingsToDefaults'))
    
    // DASHBOARDS - Multi-dimensional analytics
    .addSubMenu(ui.createMenu('📊 Dashboards')
      .addItem('🔄 Refresh All Dashboards', 'refreshAllDashboards')
      .addSeparator()
      .addItem('💰 Finance Dashboard - Costs & Billing', 'createFinanceDashboard')
      .addItem('📊 Add Charts to Finance Dashboard', 'addFinanceDashboardCharts')
      .addItem('🔄 Refresh Finance Costs → Assumptions', 'refreshFinanceDashboard')
      .addItem('📈 CEO Dashboard - Executive overview', 'runCreateCEODashboard')
      .addItem('📊 Pivot Analysis - Multi-dimensional views', 'runCreatePivotAnalysis')
      .addSeparator()
      .addItem('🏢 Partner View - Customer impact', 'generatePartnerDashboard')
      .addItem('💼 Vendor View - Cost exposure', 'generateVendorDashboard')
      .addItem('📍 State View - Jurisdiction volume', 'generateStateDashboard')
      .addItem('🗺️ Geographic View - Regional analysis', 'generateGeographicDashboard')
      .addSeparator()
      .addItem('📅 Monthly Reconciliation (Current Year)', 'runMonthlyReconciliation'))
    
    // EXPORT DATA - Download options
    .addSubMenu(ui.createMenu('📤 Export Data')
      .addItem('📄 Export Raw Tickets (CSV)', 'exportRawDataCSV')
      .addItem('📄 Export Processed History', 'exportFilteredData')
      .addItem('📄 Export Overrides Only', 'exportOverridesOnly')
      .addItem('📄 Export Change History', 'exportAuditLog'))
    
    // VIEW SHEETS - Navigation
    .addSubMenu(ui.createMenu('👁️ View Sheets')
      .addItem('📋 Raw Tickets - Source data from API', 'navigateToRaw')
      .addItem('📊 Ticket History - Enriched & classified', 'navigateToHistory')
      .addItem('📈 Executive Dashboard - Summary metrics', 'navigateToDashboard')
      .addItem('🏢 Partner Summary - By company', 'navigateToPartners'))
    .addSeparator()
    
    // SETTINGS - Configuration
    .addSubMenu(ui.createMenu('⚙️ Settings')
      .addItem('🔑 Configure Freshdesk API Key', 'setupScriptProperties')
      .addItem('🔌 Test API Connection', 'testApiConnectionMenu')
      .addSeparator()
      .addItem('⏰ Setup Auto-Refresh (Hourly + Daily)', 'setupAutomaticTriggers')
      .addItem('📋 View Active Triggers', 'viewActiveTriggers')
      .addItem('🗑️ Remove All Triggers', 'deleteAllTriggers')
      .addSeparator()
      .addItem('📚 Help & Documentation', 'showDocumentation'))
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════════
// UNIFIED PIPELINE HANDLERS (New User-Friendly Functions)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Backwards-compat wrapper: Ensure menu callback exists even if original wrapper omitted.
 * Some deployed projects reference `runCreateFinanceDashboard`; keep a thin wrapper
 * that forwards to the canonical `createFinanceDashboard` implementation.
 */
function runCreateFinanceDashboard() {
  try {
    return createFinanceDashboard();
  } catch (e) {
    Logger.log('runCreateFinanceDashboard wrapper error: ' + e.message);
    throw e;
  }
}


/**
 * QUICK REFRESH - 2 hour lookback
 * Fast update for recent ticket changes
 * Pipeline: API (2hr) → Raw → History
 */
function runQuickRefresh() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '⚡ Quick Refresh (2 Hours)',
    'This will:\n' +
    '• Pull tickets updated in the last 2 hours from Freshdesk\n' +
    '• Track changes in the Change Log column\n' +
    '• Process to History with enrichment\n\n' +
    'Takes about 1-2 minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('\n🚀 === QUICK REFRESH STARTED ===');
    
    // Step 1: Pull raw (2 hour lookback)
    const rawResult = pullRawHourly();
    if (!rawResult.success) {
      throw new Error(`Raw pull failed: ${rawResult.error}`);
    }
    
    // Step 2: Process to history
    const historyResult = processRawToHistory();
    if (!historyResult.success) {
      throw new Error(`History processing failed: ${historyResult.error}`);
    }
    
    const totalDuration = rawResult.duration + historyResult.duration;
    
    ui.alert('✅ Quick Refresh Complete', 
      `Tickets pulled: ${rawResult.totalTickets}\n` +
      `Changes detected: ${rawResult.changedTickets}\n` +
      `New to history: ${historyResult.newTickets}\n` +
      `Updated in history: ${historyResult.updatedTickets}\n` +
      `Duration: ${totalDuration.toFixed(1)}s`,
      ui.ButtonSet.OK);
    
    Logger.log('✅ === QUICK REFRESH COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Quick refresh failed: ${e.message}`);
    ui.alert('Error', `Quick refresh failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * FULL REFRESH - 30 day lookback
 * Complete data pull for comprehensive sync
 * Pipeline: API (30 days) → Raw → History
 */
function runFullRefresh() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '📦 Full Refresh (30 Days)',
    'This will:\n' +
    '• Pull ALL tickets from the last 30 days\n' +
    '• Track changes in the Change Log column\n' +
    '• Process to History with full enrichment\n\n' +
    '⚠️ Takes 5-10 minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Full refresh started... Check the execution log for progress.', ui.ButtonSet.OK);
  
  try {
    Logger.log('\n🚀 === FULL REFRESH STARTED ===');
    
    // Step 1: Pull raw (30 day lookback)
    const rawResult = pullRawFull();
    if (!rawResult.success) {
      throw new Error(`Raw pull failed: ${rawResult.error}`);
    }
    
    // Step 2: Process to history
    const historyResult = processRawToHistory();
    if (!historyResult.success) {
      throw new Error(`History processing failed: ${historyResult.error}`);
    }
    
    const totalDuration = rawResult.duration + historyResult.duration;
    
    ui.alert('✅ Full Refresh Complete', 
      `Tickets pulled: ${rawResult.totalTickets}\n` +
      `Changes detected: ${rawResult.changedTickets}\n` +
      `New to history: ${historyResult.newTickets}\n` +
      `Updated in history: ${historyResult.updatedTickets}\n` +
      `Duration: ${totalDuration.toFixed(1)}s`,
      ui.ButtonSet.OK);
    
    Logger.log('✅ === FULL REFRESH COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Full refresh failed: ${e.message}`);
    ui.alert('Error', `Full refresh failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * REPROCESS ALL - Re-enrich existing raw data
 * Forces reprocessing without new API pull
 * Pipeline: Raw (existing) → History (force update)
 */
function runReprocessAll() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🔁 Reprocess All Tickets',
    'This will:\n' +
    '• Re-enrich ALL tickets from Raw sheet\n' +
    '• Update classifications and outcomes\n' +
    '• Preserve manual overrides\n\n' +
    'Use this after updating outcome rules.\n\n' +
    '⚠️ May take several minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('\n🚀 === REPROCESS ALL STARTED ===');
    
    const result = processRawToHistory({ forceReprocess: true });
    
    if (result.success) {
      ui.alert('✅ Reprocess Complete', 
        `Tickets processed: ${result.processed}\n` +
        `Updated: ${result.updatedTickets}\n` +
        `Duration: ${result.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      throw new Error(result.error);
    }
    
    Logger.log('✅ === REPROCESS ALL COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Reprocess failed: ${e.message}`);
    ui.alert('Error', `Reprocess failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * VALIDATE DATA COLUMNS - Check if all required columns for reporting are present
 * Shows UI with validation status
 */
function runValidateDataColumns() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const validation = validateReportingColumns();
    
    let message = `📊 DATA VALIDATION REPORT\n\n`;
    message += `Sheet: ${SHEET_NAMES.MVR_TICKET_HISTORY}\n`;
    message += `Total Rows: ${validation.rowCount}\n\n`;
    
    if (validation.allPresent) {
      message += `✅ ALL REQUIRED COLUMNS PRESENT\n\n`;
    } else {
      message += `⚠️ MISSING COLUMNS DETECTED\n\n`;
    }
    
    message += `REQUIRED FOR REPORTS:\n`;
    validation.required.forEach(col => {
      const status = col.present ? '✅' : '❌';
      const fill = col.present ? `(${col.fillRate}% filled)` : 'MISSING';
      message += `${status} ${col.name} ${fill}\n`;
    });
    
    message += `\nDATA QUALITY:\n`;
    message += `• Vendor Group filled: ${validation.vendorFillRate}%\n`;
    message += `• DL State filled: ${validation.stateFillRate}%\n`;
    message += `• Partner filled: ${validation.partnerFillRate}%\n`;
    message += `• Outcome filled: ${validation.outcomeFillRate}%\n`;
    
    if (validation.vendorBreakdown) {
      message += `\nVENDOR BREAKDOWN:\n`;
      Object.entries(validation.vendorBreakdown)
        .sort((a, b) => b[1] - a[1])
        .forEach(([v, c]) => {
          message += `• ${v}: ${c} (${((c / validation.rowCount) * 100).toFixed(1)}%)\n`;
        });
    }
    
    const buttonSet = validation.allPresent ? ui.ButtonSet.OK : ui.ButtonSet.OK;
    ui.alert(validation.allPresent ? '✅ Data Ready for Reports' : '⚠️ Data Validation Issues', 
      message, buttonSet);
    
  } catch (e) {
    Logger.log(`❌ Validation failed: ${e.message}`);
    ui.alert('Error', `Validation failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Validate that all required columns for reporting exist and have data
 */
function validateReportingColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      allPresent: false,
      rowCount: 0,
      required: [],
      error: 'No ticket history found'
    };
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowCount = data.length - 1;
  
  // Build column map
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);
  
  // Required columns for dashboards and reports
  const requiredColumns = [
    'Ticket ID',
    'Email Subject Line',
    'Request Type (SC/RC)',
    'Partner Name',
    'Turn ID',
    'Status',
    'Date Created',
    'Date Resolved',
    'Resolution Time (Hours)',
    'Vendor Group',
    'DL State',
    'MVR Outcome',
    'Tags'
  ];
  
  const required = requiredColumns.map(colName => {
    const colIdx = colMap[colName];
    const present = colIdx !== undefined;
    let fillCount = 0;
    
    if (present) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][colIdx] && String(data[i][colIdx]).trim() !== '') {
          fillCount++;
        }
      }
    }
    
    return {
      name: colName,
      present,
      fillRate: present ? ((fillCount / rowCount) * 100).toFixed(1) : 0
    };
  });
  
  const allPresent = required.every(col => col.present);
  
  // Calculate specific fill rates
  const vendorCol = colMap['Vendor Group'];
  const stateCol = colMap['DL State'];
  const partnerCol = colMap['Partner Name'];
  const outcomeCol = colMap['MVR Outcome'];
  
  let vendorFill = 0, stateFill = 0, partnerFill = 0, outcomeFill = 0;
  const vendorBreakdown = {};
  
  for (let i = 1; i < data.length; i++) {
    if (vendorCol !== undefined && data[i][vendorCol]) {
      vendorFill++;
      const v = String(data[i][vendorCol]);
      vendorBreakdown[v] = (vendorBreakdown[v] || 0) + 1;
    }
    if (stateCol !== undefined && data[i][stateCol]) stateFill++;
    if (partnerCol !== undefined && data[i][partnerCol]) partnerFill++;
    if (outcomeCol !== undefined && data[i][outcomeCol]) outcomeFill++;
  }
  
  return {
    allPresent,
    rowCount,
    required,
    vendorFillRate: ((vendorFill / rowCount) * 100).toFixed(1),
    stateFillRate: ((stateFill / rowCount) * 100).toFixed(1),
    partnerFillRate: ((partnerFill / rowCount) * 100).toFixed(1),
    outcomeFillRate: ((outcomeFill / rowCount) * 100).toFixed(1),
    vendorBreakdown
  };
}

/**
 * REFRESH VENDOR GROUPS - Re-calculate Vendor Group from existing row data
 * Uses Subject Line + DL State + All Notes Text to correctly identify vendor
 * Fixes UNKNOWN values to CERTN, PENNDOT, INFORMDATA when data is available
 */
function runRefreshVendorGroups() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🏷️ Refresh Vendor Groups',
    'This will:\n' +
    '• Re-calculate Vendor Group for ALL tickets in History\n' +
    '• Use Subject Line + DL State to identify vendor\n' +
    '• Fall back to All Notes Text (Raw sheet) if DL State is missing\n' +
    '• Also update DL State column if found in notes\n\n' +
    'This is fast (no API calls).\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('\n🏷️ === REFRESH VENDOR GROUPS STARTED ===');
    const startTime = new Date();
    
    const refreshResult = refreshVendorGroupsInHistory();
    
    if (refreshResult.success) {
      const duration = (new Date() - startTime) / 1000;
      ui.alert('✅ Vendor Groups Refreshed', 
        `Tickets processed: ${refreshResult.total}\n` +
        `Updated: ${refreshResult.updated}\n` +
        `Unchanged: ${refreshResult.unchanged}\n` +
        `States from notes: ${refreshResult.stateFromNotes || 0}\n` +
        `Vendors from notes: ${refreshResult.vendorFromNotes || 0}\n\n` +
        `Changes by vendor:\n` +
        Object.entries(refreshResult.vendorCounts || {})
          .map(([v, c]) => `  ${v}: ${c}`)
          .join('\n') +
        `\n\nDuration: ${duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      throw new Error(refreshResult.error);
    }
    
    Logger.log('✅ === REFRESH VENDOR GROUPS COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Refresh vendor groups failed: ${e.message}`);
    ui.alert('Error', `Refresh failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Refresh Vendor Group column in History sheet
 * Uses extractVendorGroup(subject, dlState) on each row
 * Falls back to All Notes Text in Raw sheet if DL State is missing
 */
function refreshVendorGroupsInHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  if (!historySheet || historySheet.getLastRow() < 2) {
    return { success: false, error: 'No ticket history found' };
  }
  
  // Load history data
  const data = historySheet.getDataRange().getValues();
  const headers = data[0];
  
  // Build column map for history
  const colMap = {};
  headers.forEach((h, i) => colMap[h] = i);
  
  const ticketIdCol = colMap['Ticket ID'];
  const subjectCol = colMap['Email Subject Line'];
  const dlStateCol = colMap['DL State'];
  const vendorGroupCol = colMap['Vendor Group'];
  
  if (subjectCol === undefined || vendorGroupCol === undefined) {
    return { success: false, error: 'Required columns not found in History' };
  }
  
  // Load raw data for All Notes Text and Description fallback
  const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
  let rawNotesMap = new Map(); // ticketId -> allNotesText
  let rawDescMap = new Map();  // ticketId -> description
  
  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getDataRange().getValues();
    const rawHeaders = rawData[0];
    const rawColMap = {};
    rawHeaders.forEach((h, i) => rawColMap[h] = i);
    
    const rawIdCol = rawColMap['Ticket ID'];
    const rawNotesCol = rawColMap['All Notes Text'];
    const rawDescCol = rawColMap['Description'];
    
    if (rawIdCol !== undefined) {
      for (let i = 1; i < rawData.length; i++) {
        const ticketId = rawData[i][rawIdCol];
        if (ticketId) {
          rawNotesMap.set(ticketId, rawData[i][rawNotesCol] || '');
          rawDescMap.set(ticketId, rawData[i][rawDescCol] || '');
        }
      }
      Logger.log(`📋 Loaded ${rawNotesMap.size} tickets from Raw for notes/description lookup`);
    }
  }
  
  Logger.log(`📋 Processing ${data.length - 1} history tickets...`);
  
  let updated = 0;
  let unchanged = 0;
  let stateFromNotes = 0;
  let vendorFromNotes = 0;
  const vendorCounts = {};
  const updates = []; // Batch updates for vendor group
  const stateUpdates = []; // Also update DL State if found in notes
  
  for (let i = 1; i < data.length; i++) {
    const ticketId = data[i][ticketIdCol];
    const subject = data[i][subjectCol] || '';
    let dlState = data[i][dlStateCol] || '';
    const currentVendor = data[i][vendorGroupCol] || '';
    
    // If DL State is missing, try to get it from Raw data
    if (!dlState && rawNotesMap.has(ticketId)) {
      const notesText = rawNotesMap.get(ticketId);
      const descText = rawDescMap.get(ticketId) || '';
      
      // Try parseAllNotesText first
      const notesData = parseAllNotesText(notesText);
      
      if (notesData.dl_state) {
        dlState = notesData.dl_state;
        stateFromNotes++;
        stateUpdates.push([i + 1, dlStateCol + 1, dlState]);
      } else {
        // Try more aggressive extraction from combined text
        const combinedText = descText + ' ' + notesText;
        const extractedState = extractStateFromText(combinedText);
        if (extractedState) {
          dlState = extractedState;
          stateFromNotes++;
          stateUpdates.push([i + 1, dlStateCol + 1, dlState]);
        }
      }
      
      // Check if notes has explicit vendor
      if (notesData.vendor) {
        const notesVendor = notesData.vendor.toUpperCase();
        if (['CERTN', 'PENNDOT', 'INFORM', 'INFORMDATA'].includes(notesVendor)) {
          if (notesVendor !== currentVendor) {
            updates.push([i + 1, vendorGroupCol + 1, notesVendor]);
            updated++;
            vendorFromNotes++;
            vendorCounts[notesVendor] = (vendorCounts[notesVendor] || 0) + 1;
          } else {
            unchanged++;
          }
          continue; // Skip normal extraction since we got explicit vendor
        }
      }
    }
    
    // Calculate correct vendor group using unified function
    const correctVendor = extractVendorGroup(subject, dlState);
    
    if (correctVendor !== currentVendor) {
      updates.push([i + 1, vendorGroupCol + 1, correctVendor]);
      updated++;
      vendorCounts[correctVendor] = (vendorCounts[correctVendor] || 0) + 1;
      
      if (updated <= 10) {
        Logger.log(`   Row ${i + 1}: "${currentVendor}" → "${correctVendor}" (state: ${dlState || 'none'})`);
      }
    } else {
      unchanged++;
    }
    
    if (i % 500 === 0) {
      Logger.log(`   Progress: ${i}/${data.length - 1}`);
    }
  }
  
  // Apply state updates first
  if (stateUpdates.length > 0) {
    Logger.log(`\n📝 Updating ${stateUpdates.length} DL State values from notes...`);
    for (const [row, col, value] of stateUpdates) {
      historySheet.getRange(row, col).setValue(value);
    }
  }
  
  // Apply vendor updates
  if (updates.length > 0) {
    Logger.log(`\n📝 Applying ${updates.length} vendor updates...`);
    for (const [row, col, value] of updates) {
      historySheet.getRange(row, col).setValue(value);
    }
  }
  
  Logger.log(`\n✅ Vendor Group Refresh Complete:`);
  Logger.log(`   Total: ${data.length - 1}`);
  Logger.log(`   Updated: ${updated}`);
  Logger.log(`   Unchanged: ${unchanged}`);
  Logger.log(`   States extracted from notes: ${stateFromNotes}`);
  Logger.log(`   Vendors explicit in notes: ${vendorFromNotes}`);
  Logger.log(`   Breakdown: ${JSON.stringify(vendorCounts)}`);
  
  return {
    success: true,
    total: data.length - 1,
    updated,
    unchanged,
    stateFromNotes,
    vendorFromNotes,
    vendorCounts
  };
}

/**
 * DEBUG: Analyze why tickets are UNKNOWN
 * Samples tickets with missing vendor/state to diagnose extraction issues
 */
function debugUnknownVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  const rawSheet = ss.getSheetByName(SHEET_NAMES.MVR_RAW_TICKETS);
  
  Logger.log('\n🔍 === DEBUG UNKNOWN VENDORS ===\n');
  
  // Load history
  const historyData = historySheet.getDataRange().getValues();
  const historyHeaders = historyData[0];
  const historyColMap = {};
  historyHeaders.forEach((h, i) => historyColMap[h] = i);
  
  // Load raw for notes lookup
  const rawData = rawSheet.getDataRange().getValues();
  const rawHeaders = rawData[0];
  const rawColMap = {};
  rawHeaders.forEach((h, i) => rawColMap[h] = i);
  
  // Build raw notes map
  const rawNotesMap = new Map();
  const rawDescMap = new Map();
  for (let i = 1; i < rawData.length; i++) {
    const ticketId = rawData[i][rawColMap['Ticket ID']];
    rawNotesMap.set(ticketId, rawData[i][rawColMap['All Notes Text']] || '');
    rawDescMap.set(ticketId, rawData[i][rawColMap['Description']] || '');
  }
  
  // Find UNKNOWN tickets and sample them
  let unknownCount = 0;
  let noStateCount = 0;
  let sampleCount = 0;
  const maxSamples = 10;
  
  for (let i = 1; i < historyData.length; i++) {
    const vendor = historyData[i][historyColMap['Vendor Group']] || '';
    const dlState = historyData[i][historyColMap['DL State']] || '';
    const subject = historyData[i][historyColMap['Email Subject Line']] || '';
    const ticketId = historyData[i][historyColMap['Ticket ID']];
    
    if (vendor === 'UNKNOWN' || vendor === '' || vendor === 'LEGACY') {
      unknownCount++;
      if (!dlState) noStateCount++;
      
      // Sample up to maxSamples
      if (sampleCount < maxSamples) {
        sampleCount++;
        Logger.log(`\n--- Sample ${sampleCount}: Ticket ${ticketId} ---`);
        Logger.log(`Subject: ${subject.substring(0, 100)}`);
        Logger.log(`Current DL State: "${dlState}"`);
        Logger.log(`Current Vendor: "${vendor}"`);
        
        // Check raw data
        const notesText = rawNotesMap.get(ticketId) || '';
        const descText = rawDescMap.get(ticketId) || '';
        
        Logger.log(`\nDescription (first 500 chars):`);
        Logger.log(descText.substring(0, 500));
        
        Logger.log(`\nAll Notes (first 500 chars):`);
        Logger.log(notesText.substring(0, 500));
        
        // Try to find state with multiple patterns
        const combinedText = descText + ' ' + notesText;
        const foundState = extractStateFromText(combinedText);
        if (foundState) {
          Logger.log(`\n✅ Extracted state: "${foundState}"`);
        } else {
          Logger.log(`\n❌ No state found in text`);
        }
      }
    }
  }
  
  Logger.log(`\n\n=== SUMMARY ===`);
  Logger.log(`Total tickets: ${historyData.length - 1}`);
  Logger.log(`UNKNOWN/LEGACY vendors: ${unknownCount}`);
  Logger.log(`Missing DL State: ${noStateCount}`);
  Logger.log(`Sampled: ${sampleCount}`);
}

/**
 * Extract state from text using multiple patterns
 * Tries various formats found in ticket data
 */
function extractStateFromText(text) {
  if (!text) return null;
  
  // Patterns to try, in order of specificity
  const patterns = [
    /DL State:\s*\*?([A-Z]{2})\*?/i,           // DL State: *TN* or DL State: TN
    /State:\s*\*?([A-Z]{2})\*?/i,               // State: TN
    /Driver['']?s?\s+License\s+State:\s*([A-Z]{2})/i,  // Driver's License State: TN
    /\bState\b[:\s]+([A-Z]{2})\b/i,             // State: TN or State TN
    /License\s+State:\s*([A-Z]{2})/i,           // License State: TN
    /\|\s*([A-Z]{2})\s*\|/,                     // | TN | (table format)
    /State Vendor:\s*[^|]+\|\s*([A-Z]{2})\b/i,  // State Vendor: xxx | TN
    /MVR.*\b([A-Z]{2})\b.*Driver/i,             // MVR ... TN ... Driver
  ];
  
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      const state = match[1].toUpperCase();
      // Validate it's a real US state code
      if (isValidStateCode(state)) {
        return state;
      }
    }
  }
  
  return null;
}

/**
 * Check if a 2-letter code is a valid US state
 */
function isValidStateCode(code) {
  const validStates = [
    'AL', 'AK', 'AZ', 'AR', 'CA', 'CO', 'CT', 'DE', 'FL', 'GA',
    'HI', 'ID', 'IL', 'IN', 'IA', 'KS', 'KY', 'LA', 'ME', 'MD',
    'MA', 'MI', 'MN', 'MS', 'MO', 'MT', 'NE', 'NV', 'NH', 'NJ',
    'NM', 'NY', 'NC', 'ND', 'OH', 'OK', 'OR', 'PA', 'RI', 'SC',
    'SD', 'TN', 'TX', 'UT', 'VT', 'VA', 'WA', 'WV', 'WI', 'WY', 'DC'
  ];
  return validStates.includes(code);
}

/**
 * RE-ENRICH DATA - Re-extract data from All Notes Text
 * For tickets where subject parsing failed (InformData/Sentinel)
 * Uses notes fallback to extract partner, state, turn ID, tier, etc.
 */
function runReEnrichment() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🔄 Re-Enrich Data from Notes',
    'This will:\n' +
    '• Re-process all tickets from Raw sheet\n' +
    '• Extract missing data from All Notes Text\n' +
    '• Update Partner, State, Turn ID, Tier for InformData tickets\n' +
    '• Preserve manual overrides\n\n' +
    'Useful for fixing tickets with parsing failures.\n\n' +
    '⚠️ May take several minutes.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('\n🚀 === RE-ENRICH FROM NOTES STARTED ===');
    
    const result = processRawToHistory({ 
      forceReprocess: true,
      useNotesExtraction: true 
    });
    
    if (result.success) {
      ui.alert('✅ Re-Enrich Complete', 
        `Tickets processed: ${result.processed}\n` +
        `Updated with notes data: ${result.notesEnriched || 0}\n` +
        `Duration: ${result.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      throw new Error(result.error);
    }
    
    Logger.log('✅ === RE-ENRICH COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Re-enrich failed: ${e.message}`);
    ui.alert('Error', `Re-enrich failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * INITIAL SETUP - First-time pull with no change tracking
 * Clears existing data and does fresh pull
 * Use this for first-time setup or complete reset
 */
function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🆕 INITIAL SETUP - First Time Pull',
    '⚠️ WARNING: This will:\n\n' +
    '• CLEAR all existing Raw Tickets data\n' +
    '• CLEAR all existing Ticket History data\n' +
    '• Pull ALL tickets from last 30 days (fresh)\n' +
    '• Process to History (no change tracking)\n\n' +
    'This is for FIRST-TIME SETUP or COMPLETE RESET.\n\n' +
    '⏱️ Takes 5-10 minutes.\n\n' +
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  // Double confirmation for destructive action
  const confirm2 = ui.alert(
    '⚠️ Confirm Data Clear',
    'This will DELETE all existing ticket data.\n\n' +
    'Type YES to confirm you want to proceed.',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm2 !== ui.Button.YES) {
    ui.alert('Cancelled', 'Initial setup was cancelled.', ui.ButtonSet.OK);
    return;
  }
  
  ui.alert('Starting', 'Initial setup starting... This will take several minutes.', ui.ButtonSet.OK);
  
  try {
    Logger.log('\n🆕 === INITIAL SETUP STARTED ===');
    const startTime = new Date();
    
    // Step 1: Clear existing sheets
    Logger.log('🗑️ Step 1: Clearing existing data...');
    clearSheetData(SHEET_NAMES.MVR_RAW_TICKETS);
    clearSheetData(SHEET_NAMES.MVR_TICKET_HISTORY);
    Logger.log('   ✅ Sheets cleared');
    
    // Step 2: Pull fresh data (no change tracking)
    Logger.log('📥 Step 2: Pulling fresh data (30 days)...');
    const rawResult = pullRawInitial();
    if (!rawResult.success) {
      throw new Error(`Raw pull failed: ${rawResult.error}`);
    }
    Logger.log(`   ✅ Pulled ${rawResult.totalTickets} tickets`);
    
    // Step 3: Process to history (force all as new)
    Logger.log('⚙️ Step 3: Processing to history...');
    const historyResult = processRawToHistory({ forceReprocess: true });
    if (!historyResult.success) {
      throw new Error(`History processing failed: ${historyResult.error}`);
    }
    Logger.log(`   ✅ Processed ${historyResult.newTickets} tickets to history`);
    
    const totalDuration = (new Date() - startTime) / 1000;
    
    ui.alert('✅ Initial Setup Complete!', 
      `🆕 Fresh data loaded successfully!\n\n` +
      `📥 Tickets pulled: ${rawResult.totalTickets}\n` +
      `📊 Tickets in history: ${historyResult.newTickets}\n\n` +
      `⏱️ Duration: ${totalDuration.toFixed(1)}s\n\n` +
      `Next steps:\n` +
      `• Run 'Generate All Reports' to create reports\n` +
      `• Setup automatic triggers for ongoing updates`,
      ui.ButtonSet.OK);
    
    Logger.log('✅ === INITIAL SETUP COMPLETE ===\n');
    
  } catch (e) {
    Logger.log(`❌ Initial setup failed: ${e.message}`);
    Logger.log(e.stack);
    ui.alert('Error', `Initial setup failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Clear all data from a sheet (keep headers)
 */
function clearSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
    Logger.log(`   Cleared ${sheetName}`);
  }
}

/**
 * COMPLETE REFRESH - Full pipeline with reports
 * One-click complete workflow
 * Pipeline: API (30 days) → Raw → History → All Reports
 */
function runCompleteRefresh() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '▶️ Complete Refresh Starting',
    'This will:\n' +
    '1. Pull all tickets (30 days)\n' +
    '2. Process to History with enrichment\n' +
    '3. Generate all reports\n\n' +
    '⏱️ Takes 10-15 minutes.\n\n' +
    'No further interaction needed.',
    ui.ButtonSet.OK
  );
  
  try {
    Logger.log('\n🚀 === COMPLETE REFRESH STARTED ===');
    const startTime = new Date();
    
    // Step 1: Full raw pull
    Logger.log('📥 Step 1: Pulling raw tickets (30 days)...');
    const rawResult = pullRawFull();
    if (!rawResult.success) {
      throw new Error(`Raw pull failed: ${rawResult.error}`);
    }
    Logger.log(`   ✅ Pulled ${rawResult.totalTickets} tickets`);
    
    // Step 2: Process to history
    Logger.log('⚙️ Step 2: Processing to history...');
    const historyResult = processRawToHistory();
    if (!historyResult.success) {
      throw new Error(`History processing failed: ${historyResult.error}`);
    }
    Logger.log(`   ✅ Processed: ${historyResult.newTickets} new, ${historyResult.updatedTickets} updated`);
    
    // Step 3: Generate reports
    Logger.log('📊 Step 3: Generating reports...');
    generatePartnerSummary();
    generateMonthlyMVRReport();
    generateExecutiveDashboard();
    generateFinanceAuditTrail();
    generateOpsPerformance();
    Logger.log('   ✅ All reports generated');
    
    const totalDuration = (new Date() - startTime) / 1000;
    
    ui.alert('✅ Complete Refresh Finished!', 
      `📥 Tickets pulled: ${rawResult.totalTickets}\n` +
      `⚙️ New to history: ${historyResult.newTickets}\n` +
      `🔄 Updated: ${historyResult.updatedTickets}\n` +
      `📊 Reports: 5 generated\n\n` +
      `⏱️ Total time: ${totalDuration.toFixed(1)}s`,
      ui.ButtonSet.OK);
    
    Logger.log('✅ === COMPLETE REFRESH FINISHED ===\n');
    
  } catch (e) {
    Logger.log(`❌ Complete refresh failed: ${e.message}`);
    ui.alert('Error', `Complete refresh failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * GENERATE ALL REPORTS - Report generation only
 */
function runAllReports() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '📊 Generate All Reports',
    'This will generate:\n' +
    '• Partner Summary\n' +
    '• Monthly MVR Report\n' +
    '• Executive Dashboard\n' +
    '• Finance Audit Trail\n' +
    '• Team Performance\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('\n📊 === GENERATING ALL REPORTS ===');
    const startTime = new Date();
    
    generatePartnerSummary();
    generateMonthlyMVRReport();
    generateExecutiveDashboard();
    generateFinanceAuditTrail();
    generateOpsPerformance();
    
    const duration = (new Date() - startTime) / 1000;
    
    ui.alert('✅ Reports Generated!', 
      `5 reports created successfully.\n\nDuration: ${duration.toFixed(1)}s`,
      ui.ButtonSet.OK);
    
    Logger.log('✅ === ALL REPORTS GENERATED ===\n');
    
  } catch (e) {
    Logger.log(`❌ Report generation failed: ${e.message}`);
    ui.alert('Error', `Report generation failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Generate Partner Summary report
 */
function runPartnerSummary() {
  try {
    generatePartnerSummary();
    SpreadsheetApp.getUi().alert('✅ Partner Summary generated!', '', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate Assumptions Log sheet
 */
function runAssumptionsLog() {
  try {
    generateAssumptionsLog();
    SpreadsheetApp.getUi().alert('✅ Assumptions Log generated!', 'See Assumptions_Log sheet for all documented assumptions.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate OPS Lifecycle Report
 */
function runOpsLifecycleReport() {
  try {
    const result = generateOpsLifecycleReport();
    SpreadsheetApp.getUi().alert('✅ OPS Lifecycle Report generated!', 
      `Analyzed ${result.totalTickets} tickets.\n${result.blockedTickets} blocked tickets identified.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate Finance Reconciliation Report
 */
function runFinanceReconciliationReport() {
  try {
    const result = generateFinanceReconciliationReport();
    SpreadsheetApp.getUi().alert('✅ Finance Reconciliation Report generated!', 
      `${result.totalChecks} completed checks analyzed.\n${result.freeChecks} FREE, ${result.paidChecks} PAID.\nSee Turn_ID_Billing_List for vendor invoicing.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate CEO Effectiveness Report
 */
function runCEOEffectivenessReport() {
  try {
    const result = generateCEOEffectivenessReport();
    SpreadsheetApp.getUi().alert('✅ CEO Effectiveness Report generated!', 
      `Detection Rate: ${result.detectionRate}%\nAvg Resolution: ${result.avgResolutionHours} hours`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate Annual Projection Report
 */
function runAnnualProjection() {
  try {
    const result = generateAnnualProjection();
    SpreadsheetApp.getUi().alert('✅ Annual Projection generated!', 
      `Projected for ${result.year}.\nSee Annual_Projection sheet for Low/Mid/High scenarios.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate Operative Plan
 */
function runOperativePlan() {
  try {
    const result = generateOperativePlan();
    SpreadsheetApp.getUi().alert('✅ Operative Plan generated!', 
      `Wave calendar for ${result.year}.\nSee Operative_Plan sheet for monthly schedule.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Generate Growth Projection Report
 */
function runGrowthProjection() {
  try {
    const result = generateGrowthProjection();
    SpreadsheetApp.getUi().alert('✅ Growth Projection generated!', 
      `5-year projection: ${result.years}.\nSee Growth_Projection sheet for hiring triggers.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error', e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// SHEET NAVIGATION HANDLERS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Navigate to Raw Tickets sheet
 */
function navigateToRaw() {
  navigateToSheet(SHEET_NAMES.MVR_RAW_TICKETS);
}

/**
 * Navigate to Ticket History sheet
 */
function navigateToHistory() {
  navigateToSheet(SHEET_NAMES.MVR_TICKET_HISTORY);
}

/**
 * Navigate to Executive Dashboard sheet
 */
function navigateToDashboard() {
  navigateToSheet('Executive_Dashboard');
}

/**
 * Navigate to Partner Summary sheet
 */
function navigateToPartners() {
  navigateToSheet('By_Partner_Summary');
}

/**
 * Navigate to a specific sheet by name
 */
function navigateToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    ss.setActiveSheet(sheet);
    SpreadsheetApp.flush();
  } else {
    SpreadsheetApp.getUi().alert('Sheet Not Found', 
      `Sheet "${sheetName}" does not exist yet.\n\nRun a refresh first to create it.`,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// AUTOMATED TRIGGER FUNCTIONS (Called by triggers, not menu)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * HOURLY TRIGGER - Runs every hour
 * Pipeline: API (2hr) → Raw → History → Refresh Dashboards
 */
function hourlyRefresh() {
  Logger.log('\n⏰ === HOURLY TRIGGER STARTED ===');
  
  try {
    // Pull raw (2 hour lookback)
    const rawResult = pullRawHourly();
    if (!rawResult.success) {
      Logger.log(`❌ Hourly raw pull failed: ${rawResult.error}`);
      return;
    }
    
    // Process to history
    const historyResult = processRawToHistory();
    if (!historyResult.success) {
      Logger.log(`❌ Hourly history processing failed: ${historyResult.error}`);
      return;
    }
    
    // Refresh dashboards (force formula recalculation)
    SpreadsheetApp.flush();
    
    Logger.log(`✅ Hourly refresh complete: ${rawResult.totalTickets} pulled, ${historyResult.newTickets} new, ${historyResult.updatedTickets} updated`);
    
  } catch (e) {
    Logger.log(`❌ Hourly refresh error: ${e.message}`);
  }
  
  Logger.log('⏰ === HOURLY TRIGGER COMPLETE ===\n');
}

/**
 * DAILY TRIGGER - Runs at 9 AM
 * Pipeline: API (30 days) → Raw → History → Reports → Dashboards → Pivots
 */
function dailyFullRefresh() {
  Logger.log('\n🌅 === DAILY 9 AM TRIGGER STARTED ===');
  
  try {
    // Step 1: Full raw pull
    Logger.log('📥 Step 1: Full raw pull (30 days)...');
    const rawResult = pullRawFull();
    if (!rawResult.success) {
      Logger.log(`❌ Daily raw pull failed: ${rawResult.error}`);
      return;
    }
    Logger.log(`   ✅ Pulled ${rawResult.totalTickets} tickets`);
    
    // Step 2: Process to history
    Logger.log('⚙️ Step 2: Processing to history...');
    const historyResult = processRawToHistory();
    if (!historyResult.success) {
      Logger.log(`❌ Daily history processing failed: ${historyResult.error}`);
      return;
    }
    Logger.log(`   ✅ ${historyResult.newTickets} new, ${historyResult.updatedTickets} updated`);
    
    // Step 3: Generate legacy reports
    Logger.log('📊 Step 3: Generating legacy reports...');
    generatePartnerSummary();
    generateMonthlyMVRReport();
    generateExecutiveDashboard();
    generateFinanceAuditTrail();
    generateOpsPerformance();
    Logger.log('   ✅ Legacy reports generated');
    
    // Step 4: Refresh live dashboards (force formula recalculation)
    Logger.log('📈 Step 4: Refreshing live dashboards...');
    refreshLiveDashboards();
    Logger.log('   ✅ Dashboards refreshed');
    
    // Step 5: Refresh pivots
    Logger.log('📊 Step 5: Refreshing pivot tables...');
    refreshPivots();
    Logger.log('   ✅ Pivots refreshed');
    
    Logger.log('✅ Daily refresh complete!');
    
  } catch (e) {
    Logger.log(`❌ Daily refresh error: ${e.message}`);
    Logger.log(e.stack);
  }
  
  Logger.log('🌅 === DAILY 9 AM TRIGGER COMPLETE ===\n');
}

/**
 * Refresh all live dashboards by forcing formula recalculation
 * Called after data refresh to ensure dashboards show latest data
 */
function refreshLiveDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // List of live dashboard sheets
  const dashboardSheets = [
    'Assumptions',
    'Finance Dashboard', 
    'CEO Dashboard',
    'Pivot Analysis'
  ];
  
  let refreshed = 0;
  
  dashboardSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      // Force recalculation by flushing
      SpreadsheetApp.flush();
      refreshed++;
      Logger.log(`   Refreshed: ${sheetName}`);
    }
  });
  
  // Update last refresh timestamp in Assumptions if it exists
  const assumptions = ss.getSheetByName('Assumptions');
  if (assumptions) {
    // Find the "Last Updated" cell and update it
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
  
  Logger.log(`✅ Refreshed ${refreshed} dashboard sheets`);
  return refreshed;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN EXECUTION FUNCTIONS (Menu Actions)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Quick sync - 10 day lookback, for periodic updates
 */
function runQuickSync() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Quick Sync (10 Days)',
    'This will:\n' +
    '• Fetch MVR tickets from the last 10 days\n' +
    '• Append NEW tickets to history\n' +
    '• UPDATE tickets where tags/status changed\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Syncing tickets... This may take a few minutes.', ui.ButtonSet.OK);
  
  try {
    const summary = syncTicketsPeriodic();
    
    if (summary.success) {
      ui.alert('Success', 
        `✅ Quick Sync Complete!\n\n` +
        `Fetched: ${summary.fetched} tickets\n` +
        `New (appended): ${summary.newTickets}\n` +
        `Changed (updated): ${summary.updatedTickets}\n` +
        `Unchanged: ${summary.unchangedTickets}\n` +
        `Duration: ${summary.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Sync failed: ${summary.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Sync failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Full sync - 30 day lookback, for manual/initial loads
 */
function runFullSync() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Full Sync (30 Days)',
    'This will:\n' +
    '• Fetch MVR tickets from the last 30 days\n' +
    '• Append NEW tickets to history\n' +
    '• UPDATE tickets where tags/status changed\n' +
    '• Fetch conversations for classification\n\n' +
    '⚠️ This may take 5-10 minutes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Full sync started... Check execution logs for progress.', ui.ButtonSet.OK);
  
  try {
    const summary = syncTicketsFull();
    
    if (summary.success) {
      ui.alert('Success', 
        `✅ Full Sync Complete!\n\n` +
        `Fetched: ${summary.fetched} tickets\n` +
        `New (appended): ${summary.newTickets}\n` +
        `Changed (updated): ${summary.updatedTickets}\n` +
        `Unchanged: ${summary.unchangedTickets}\n` +
        `Duration: ${summary.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Sync failed: ${summary.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Sync failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Raw Hourly Sync - 2 hour lookback with change tracking
 * Updates raw tickets and logs changes inline
 */
function runRawHourlySync() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Raw Hourly Sync (2 Hours)',
    'This will:\n' +
    '• Fetch MVR tickets updated in the last 2 hours\n' +
    '• Compare with existing raw data\n' +
    '• Log changes inline (Change Log column)\n' +
    '• Overwrite MVR_Raw_Tickets with fresh data\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    const summary = pullRawHourly();
    
    if (summary.success) {
      ui.alert('Success', 
        `✅ Hourly Sync Complete!\n\n` +
        `Total tickets: ${summary.totalTickets}\n` +
        `New: ${summary.newTickets}\n` +
        `Changed: ${summary.changedTickets}\n` +
        `Duration: ${summary.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Sync failed: ${summary.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Sync failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Raw Full Refresh - 30 day lookback with change tracking
 * Complete refresh of raw tickets with change history preserved
 */
function runRawFullRefresh() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Raw Full Refresh (30 Days)',
    'This will:\n' +
    '• Fetch ALL MVR tickets from the last 30 days\n' +
    '• Compare with existing raw data\n' +
    '• Log changes inline (Change Log column)\n' +
    '• Overwrite MVR_Raw_Tickets with fresh data\n\n' +
    '⚠️ This may take several minutes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Full refresh started... This may take a few minutes.', ui.ButtonSet.OK);
  
  try {
    const summary = pullRawFull();
    
    if (summary.success) {
      ui.alert('Success', 
        `✅ Full Refresh Complete!\n\n` +
        `Total tickets: ${summary.totalTickets}\n` +
        `New: ${summary.newTickets}\n` +
        `Changed: ${summary.changedTickets}\n` +
        `Duration: ${summary.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Refresh failed: ${summary.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Refresh failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Full update with sync - runs full sync then generates reports
 */
function runFullUpdateWithSync() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Full Update (Sync + Reports)',
    'This will:\n' +
    '• Full Sync (30 days) - append new, update changed\n' +
    '• Generate all reports\n\n' +
    '⚠️ This may take 10-15 minutes.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Starting full update... Check execution logs for progress.', ui.ButtonSet.OK);
  
  try {
    // Run full sync
    const syncResult = syncTicketsFull();
    
    if (!syncResult.success) {
      ui.alert('Error', `Sync failed: ${syncResult.error}`, ui.ButtonSet.OK);
      return;
    }
    
    // Generate reports
    generatePartnerSummary();
    generateMonthlyMVRReport();
    generateExecutiveDashboard();
    generateFinanceAuditTrail();
    generateOpsPerformance();
    
    ui.alert('Success', 
      `✅ Full Update Complete!\n\n` +
      `Synced: ${syncResult.fetched} tickets\n` +
      `New: ${syncResult.newTickets}\n` +
      `Updated: ${syncResult.updatedTickets}\n` +
      `Reports generated: 5\n` +
      `Duration: ${syncResult.duration.toFixed(1)}s`,
      ui.ButtonSet.OK);
      
  } catch (e) {
    ui.alert('Error', `Update failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Called by time-based trigger (hourly/daily)
 */
function runScheduledSync() {
  Logger.log('⏰ Running scheduled sync (10-day lookback)...');
  const result = syncTicketsPeriodic();
  Logger.log(`Scheduled sync complete: ${result.newTickets} new, ${result.updatedTickets} updated`);
  return result;
}

/**
 * Run fetch and append from menu (legacy - redirects to quick sync)
 */
function runFetchAndAppend() {
  runQuickSync();
}

/**
 * Run update existing tickets from menu
 */
function runUpdateExisting() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Update Existing Tickets',
    'This will refresh status/resolution times for all open/pending tickets.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  ui.alert('Processing', 'Updating tickets... This may take several minutes.', ui.ButtonSet.OK);
  
  try {
    const count = updateExistingTickets();
    ui.alert('Success', `✅ Updated ${count} tickets in history`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Update failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run generate reports from menu
 */
function runGenerateReports() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Generate Reports',
    'This will generate partner summary reports from historical data.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  try {
    const summary = generateAllReports();
    
    if (summary.success) {
      const message = `✅ Reports Generated!\n\n` +
        `Partners: ${summary.partners}\n` +
        `Duration: ${summary.duration.toFixed(1)}s`;
      
      ui.alert('Success', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Report generation failed: ${summary.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Report generation failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run complete flow - ONE-CLICK BUTTON
 * Executes entire workflow: fetch → append → analyze → report
 * NO user interaction required after clicking
 */
function runCompleteFlow() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Starting Complete Flow',
    'This will:\n' +
    '1. Fetch new tickets (10-day lookback)\n' +
    '2. Append to history\n' +
    '3. Generate partner reports\n' +
    '4. Generate monthly reports\n\n' +
    'No further interaction needed.\n' +
    'This may take 5-10 minutes.',
    ui.ButtonSet.OK
  );
  
  try {
    // Step 1: Fetch and append
    Logger.log('🚀 Step 1: Fetching and appending tickets...');
    const fetchSummary = fetchAndAppendNewTickets();
    
    if (!fetchSummary.success) {
      throw new Error(`Fetch failed: ${fetchSummary.error}`);
    }
    
    // Step 2: Generate partner reports
    Logger.log('📊 Step 2: Generating partner reports...');
    const partnerSummary = generatePartnerSummary();
    writePartnerSummary(partnerSummary);
    
    // Step 3: Generate monthly reports
    Logger.log('📅 Step 3: Generating monthly reports...');
    const monthlyReport = generateMonthlyRequestReport();
    writeMonthlyReport(monthlyReport);
    
    // Success message
    const message = `✅ COMPLETE FLOW FINISHED!\n\n` +
      `Fetched: ${fetchSummary.fetched} tickets\n` +
      `New: ${fetchSummary.new} tickets\n` +
      `Appended: ${fetchSummary.appended} tickets\n` +
      `Partners (By_Partner): ${partnerSummary.total}\n` +
      `Partner-Months (Monthly): ${monthlyReport.total}\n\n` +
      `Check these sheets:\n` +
      `• MVR_Ticket_History\n` +
      `• By_Partner_Summary\n` +
      `• Monthly_MVR_Report (by partner with requested Turn IDs)`;
    
    ui.alert('✅ Success', message, ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Complete flow failed: ${e.message}`);
    ui.alert('Error', `Complete flow failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run full update (fetch + reports) from menu
 */
function runFullUpdate() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Full Update',
    'This will:\n1. Fetch new tickets (10-day lookback)\n2. Append to history\n3. Generate reports\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  ui.alert('Processing', 'Running full update... This may take several minutes.', ui.ButtonSet.OK);
  
  try {
    // Step 1: Fetch and append
    const fetchSummary = fetchAndAppendNewTickets();
    
    if (!fetchSummary.success) {
      ui.alert('Error', `Fetch failed: ${fetchSummary.error}`, ui.ButtonSet.OK);
      return;
    }
    
    // Step 2: Generate reports
    const reportSummary = generateAllReports();
    
    if (!reportSummary.success) {
      ui.alert('Warning', `Tickets fetched successfully but report generation failed: ${reportSummary.error}`, ui.ButtonSet.OK);
      return;
    }
    
    // Success
    const message = `✅ Full Update Complete!\n\n` +
      `Fetched: ${fetchSummary.fetched} tickets\n` +
      `New: ${fetchSummary.new} tickets\n` +
      `Appended: ${fetchSummary.appended} tickets\n` +
      `Partners: ${reportSummary.partners}\n` +
      `Partner-Months: ${reportSummary.months || 0}\n` +
      `Total Duration: ${(fetchSummary.duration + reportSummary.duration).toFixed(1)}s`;
    
    ui.alert('Success', message, ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Full update failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// TRIGGER SETUP
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Setup automatic triggers for unified pipeline
 * - Hourly: Raw (2hr) + History processing
 * - Daily 9 AM: Full refresh (30 days) + History + Reports
 */
function setupAutomaticTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '⏰ Setup Automatic Triggers',
    'This will create:\n\n' +
    '🔄 HOURLY TRIGGER:\n' +
    '   • Pull tickets updated in last 2 hours\n' +
    '   • Process to History sheet\n\n' +
    '🌅 DAILY TRIGGER (9:00 AM):\n' +
    '   • Full refresh (30 days)\n' +
    '   • Process to History\n' +
    '   • Generate all reports\n\n' +
    'Existing MVR triggers will be removed first.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    // Delete existing MVR triggers first
    deleteAllMVRTriggers();
    
    // Create HOURLY trigger
    ScriptApp.newTrigger('hourlyRefresh')
      .timeBased()
      .everyHours(1)
      .create();
    Logger.log('✅ Created hourly trigger for hourlyRefresh()');
    
    // Create DAILY 9 AM trigger
    ScriptApp.newTrigger('dailyFullRefresh')
      .timeBased()
      .atHour(9)
      .everyDays(1)
      .inTimezone(TIME_CONFIG.TIMEZONE)
      .create();
    Logger.log('✅ Created daily 9 AM trigger for dailyFullRefresh()');
    
    ui.alert('✅ Triggers Created!', 
      '⏰ Hourly Trigger:\n' +
      '   • Function: hourlyRefresh()\n' +
      '   • Schedule: Every hour\n\n' +
      '🌅 Daily Trigger:\n' +
      '   • Function: dailyFullRefresh()\n' +
      '   • Schedule: 9:00 AM daily\n' +
      '   • Timezone: ' + TIME_CONFIG.TIMEZONE,
      ui.ButtonSet.OK);
    
  } catch (e) {
    Logger.log(`❌ Error setting up triggers: ${e.message}`);
    ui.alert('Error', `Failed to setup triggers: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Daily update function (called by trigger at 9 AM)
 * Fetches new tickets and generates all reports
 */
function dailyUpdate() {
  Logger.log('\n🕐 === DAILY UPDATE STARTED (9:00 AM) ===\n');
  
  try {
    // Step 1: Fetch and append new tickets
    Logger.log('Step 1: Fetching and appending tickets...');
    const fetchSummary = fetchAndAppendNewTickets();
    
    if (!fetchSummary.success) {
      Logger.log(`❌ Daily update failed at fetch stage: ${fetchSummary.error}`);
      sendErrorNotification('Daily Update Failed', `Fetch stage failed: ${fetchSummary.error}`);
      return;
    }
    
    // Step 2: Generate partner reports
    Logger.log('Step 2: Generating partner reports...');
    const partnerSummary = generatePartnerSummary();
    writePartnerSummary(partnerSummary);
    
    // Step 3: Generate monthly reports
    Logger.log('Step 3: Generating monthly reports...');
    const monthlyReport = generateMonthlyRequestReport();
    writeMonthlyReport(monthlyReport);
    
    Logger.log('\n✅ === DAILY UPDATE COMPLETE ===');
    Logger.log(`   Fetched: ${fetchSummary.fetched} | New: ${fetchSummary.new} | Appended: ${fetchSummary.appended}`);
    Logger.log(`   Partners: ${partnerSummary.total}`);
    Logger.log(`   Partner-Months: ${monthlyReport.total}`);
    Logger.log(`   Total Duration: ${fetchSummary.duration.toFixed(1)}s\n`);
    
  } catch (e) {
    Logger.log(`❌ Daily update failed with exception: ${e.message}`);
    Logger.log(e.stack);
    sendErrorNotification('Daily Update Exception', e.message);
  }
}

/**
 * Send error notification email (if configured)
 * @param {string} subject - Email subject
 * @param {string} message - Error message
 */
function sendErrorNotification(subject, message) {
  try {
    const email = Session.getActiveUser().getEmail();
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: `MVR Tracker: ${subject}`,
        body: `Error in MVR Ticket Tracker:\n\n${message}\n\nCheck execution logs for details.`
      });
      Logger.log(`📧 Error notification sent to ${email}`);
    }
  } catch (e) {
    Logger.log(`⚠️ Could not send error notification: ${e.message}`);
  }
}

/**
 * Delete all MVR Tracker triggers
 */
function deleteAllMVRTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyUpdate') {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  Logger.log(`🗑️ Deleted ${count} MVR Tracker triggers`);
  return count;
}

/**
 * Delete all triggers (menu action)
 */
function deleteAllTriggers() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Delete Triggers',
    'This will delete all automatic triggers.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  const count = deleteAllMVRTriggers();
  ui.alert('Success', `✅ Deleted ${count} trigger(s)`, ui.ButtonSet.OK);
}

/**
 * View active triggers
 */
function viewActiveTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  const mvrTriggers = triggers.filter(t => t.getHandlerFunction() === 'dailyUpdate');
  
  if (mvrTriggers.length === 0) {
    SpreadsheetApp.getUi().alert('Active Triggers', 'No MVR Tracker triggers found.\n\nUse Setup → Setup Automatic Triggers to create one.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const triggerInfo = mvrTriggers.map(t => {
    const eventType = t.getEventType();
    return `Function: ${t.getHandlerFunction()}\nType: ${eventType}`;
  }).join('\n\n');
  
  SpreadsheetApp.getUi().alert('Active Triggers', `Found ${mvrTriggers.length} trigger(s):\n\n${triggerInfo}`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MENU HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Test API connection from menu
 */
function testApiConnectionMenu() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const creds = getFreshdeskCredentials();
    const success = testApiConnection(creds.apiKey, creds.domain);
    
    if (success) {
      ui.alert('API Test', '✅ Freshdesk API connection successful!', ui.ButtonSet.OK);
    } else {
      ui.alert('API Test', '❌ Freshdesk API connection failed. Check logs for details.', ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('API Test', `❌ Error: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Test MVR ticket fetch from menu
 */
function testMVRTicketFetchMenu() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert('Test Fetch', 'Testing MVR ticket fetch (last 24 hours)...\n\nCheck logs for results.', ui.ButtonSet.OK);
  
  try {
    const count = testMVRTicketFetch();
    ui.alert('Test Results', `✅ Test complete!\n\nFound ${count} MVR tickets in last 24 hours.\n\nCheck logs for details.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Test Results', `❌ Test failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Show history stats from menu
 */
function showHistoryStatsMenu() {
  const stats = getHistoryStats();
  
  if (stats.error) {
    SpreadsheetApp.getUi().alert('Error', stats.error, SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const message = `
📊 HISTORY STATISTICS

Total Tickets: ${stats.total}

By Type:
  • SC: ${stats.by_type.SC || 0}
  • RC: ${stats.by_type.RC || 0}

By Status:
  • Open: ${stats.by_status.Open || 0}
  • Pending: ${stats.by_status.Pending || 0}
  • Resolved: ${stats.by_status.Resolved || 0}
  • Closed: ${stats.by_status.Closed || 0}

Date Range:
  ${stats.oldest_created?.toLocaleDateString()} to ${stats.newest_created?.toLocaleDateString()}

Unique Partners: ${Object.keys(stats.by_partner || {}).length}
  `;
  
  SpreadsheetApp.getUi().alert('History Statistics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * View partner summary sheet
 */
function viewPartnerSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.BY_PARTNER_SUMMARY);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Partner Summary', 'Partner summary not found.\n\nRun "Generate Reports" first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  sheet.activate();
}

// ═══════════════════════════════════════════════════════════════════════════════
// NEW: STAKEHOLDER REPORT MENU ACTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Run Executive Dashboard from menu
 */
function runExecutiveDashboard() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Processing', 'Generating Executive Dashboard...', ui.ButtonSet.OK);
    const result = generateExecutiveDashboard();
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.EXECUTIVE_DASHBOARD);
    if (sheet) sheet.activate();
    
    ui.alert('Success', '✅ Executive Dashboard generated!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to generate dashboard: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run Finance Audit Trail from menu
 */
function runFinanceAuditTrail() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Processing', 'Generating Finance Audit Trail...', ui.ButtonSet.OK);
    const result = generateFinanceAuditTrail();
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.FINANCE_AUDIT_TRAIL);
    if (sheet) sheet.activate();
    
    ui.alert('Success', `✅ Finance Audit Trail generated!\n\nTotal records: ${result.total}`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to generate audit trail: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run Ops Performance from menu
 */
function runOpsPerformance() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    ui.alert('Processing', 'Generating Ops Performance Report...', ui.ButtonSet.OK);
    const result = generateOpsPerformance();
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.OPS_PERFORMANCE);
    if (sheet) sheet.activate();
    
    ui.alert('Success', `✅ Ops Performance Report generated!\n\nAgents: ${result.agents}`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to generate ops performance: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Run all stakeholder reports
 */
function runAllStakeholderReports() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Generate All Stakeholder Reports',
    'This will generate:\n' +
    '• Executive Dashboard\n' +
    '• Finance Audit Trail\n' +
    '• Ops Performance Report\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  try {
    Logger.log('📊 Generating all stakeholder reports...');
    
    generateExecutiveDashboard();
    generateFinanceAuditTrail();
    generateOpsPerformance();
    
    ui.alert('Success', '✅ All stakeholder reports generated!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to generate reports: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * View Audit Log sheet
 */
function viewAuditLog() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.OVERRIDE_AUDIT_LOG);
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Audit Log', 'Override Audit Log not found.\n\nNo overrides have been logged yet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  sheet.activate();
}

/**
 * Refresh override highlighting on history sheet
 */
function refreshOverrideHighlighting() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      ui.alert('Error', 'History sheet not found.', ui.ButtonSet.OK);
      return;
    }
    
    applyOverrideHighlighting(historySheet);
    ui.alert('Success', '🎨 Override highlighting refreshed!', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to refresh highlighting: ${e.message}`, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// NEW: BACKFILL MENU ACTION
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Run historical backfill from menu
 */
function runBackfillHistoricalData() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '🔄 Backfill Historical Data',
    'This will update ALL existing tickets with:\n' +
    '• Conversation data\n' +
    '• Outcome classification\n' +
    '• Workflow origin fields (DL State, Tier, etc.)\n' +
    '• Vendor group detection\n\n' +
    '⚠️ This may take 10-30 minutes depending on ticket count.\n' +
    'Processing: 50 tickets per batch with 1-second delays.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) return;
  
  ui.alert('Processing', 'Backfill started. Check execution logs for progress.\n\nThis will take several minutes.', ui.ButtonSet.OK);
  
  try {
    const backfillResult = backfillAllTickets();
    
    if (backfillResult.success) {
      ui.alert('Success', 
        `✅ Backfill Complete!\n\n` +
        `Total tickets: ${backfillResult.total}\n` +
        `Retrieved: ${backfillResult.retrieved}\n` +
        `Updated: ${backfillResult.updated}\n` +
        `Duration: ${backfillResult.duration.toFixed(1)}s`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('Error', `Backfill failed: ${backfillResult.error}`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', `Backfill failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// NEW: DATA EXPORT FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Export raw data to CSV in Google Drive
 */
function exportRawDataCSV() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    
    if (!historySheet) {
      ui.alert('Error', 'History sheet not found.', ui.ButtonSet.OK);
      return;
    }
    
    const data = historySheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('Error', 'No data to export.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert to CSV
    const csv = data.map(row => 
      row.map(cell => {
        const str = String(cell);
        // Escape quotes and wrap in quotes if contains comma
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');
    
    // Create file in Drive
    const timestamp = Utilities.formatDate(new Date(), TIME_CONFIG.TIMEZONE, "yyyyMMdd_HHmmss");
    const filename = `MVR_Ticket_Export_${timestamp}.csv`;
    
    const file = DriveApp.createFile(filename, csv, MimeType.CSV);
    
    ui.alert('Success', 
      `✅ Export Complete!\n\n` +
      `File: ${filename}\n` +
      `Rows: ${data.length - 1}\n` +
      `Location: Google Drive root folder\n\n` +
      `File URL: ${file.getUrl()}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Export failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Export filtered data based on user input
 */
function exportFilteredData() {
  const ui = SpreadsheetApp.getUi();
  
  // Get date range from user
  const startDateResp = ui.prompt('Export Filtered Data', 'Enter start date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (startDateResp.getSelectedButton() !== ui.Button.OK) return;
  
  const endDateResp = ui.prompt('Export Filtered Data', 'Enter end date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (endDateResp.getSelectedButton() !== ui.Button.OK) return;
  
  const partnerResp = ui.prompt('Export Filtered Data', 'Enter partner name (leave blank for all):', ui.ButtonSet.OK_CANCEL);
  if (partnerResp.getSelectedButton() !== ui.Button.OK) return;
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    if (!historySheet) {
      ui.alert('Error', 'History sheet not found.', ui.ButtonSet.OK);
      return;
    }
    
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    
    const startDate = new Date(startDateResp.getResponseText());
    const endDate = new Date(endDateResp.getResponseText());
    const partnerFilter = partnerResp.getResponseText().trim().toLowerCase();
    
    // Filter data
    const filtered = [headers];
    for (let i = 1; i < data.length; i++) {
      const created = new Date(data[i][colMap['Date Created']]);
      const partner = (data[i][colMap['Partner Name']] || '').toLowerCase();
      
      if (created >= startDate && created <= endDate) {
        if (!partnerFilter || partner.includes(partnerFilter)) {
          filtered.push(data[i]);
        }
      }
    }
    
    if (filtered.length <= 1) {
      ui.alert('No Data', 'No tickets match the filter criteria.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert to CSV
    const csv = filtered.map(row => 
      row.map(cell => {
        const str = String(cell);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');
    
    const timestamp = Utilities.formatDate(new Date(), TIME_CONFIG.TIMEZONE, "yyyyMMdd_HHmmss");
    const filename = `MVR_Filtered_Export_${timestamp}.csv`;
    
    const file = DriveApp.createFile(filename, csv, MimeType.CSV);
    
    ui.alert('Success', 
      `✅ Filtered Export Complete!\n\n` +
      `File: ${filename}\n` +
      `Rows: ${filtered.length - 1}\n` +
      `File URL: ${file.getUrl()}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Export failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Export only tickets with manual overrides
 */
function exportOverridesOnly() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
    if (!historySheet) {
      ui.alert('Error', 'History sheet not found.', ui.ButtonSet.OK);
      return;
    }
    
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = createColumnMapper(headers);
    const overrideIdx = colMap['Outcome Override'];
    
    if (overrideIdx === undefined) {
      ui.alert('Error', 'Outcome Override column not found.', ui.ButtonSet.OK);
      return;
    }
    
    // Filter for overrides only
    const filtered = [headers];
    for (let i = 1; i < data.length; i++) {
      if (data[i][overrideIdx]) {
        filtered.push(data[i]);
      }
    }
    
    if (filtered.length <= 1) {
      ui.alert('No Data', 'No tickets have manual overrides.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert to CSV
    const csv = filtered.map(row => 
      row.map(cell => {
        const str = String(cell);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');
    
    const timestamp = Utilities.formatDate(new Date(), TIME_CONFIG.TIMEZONE, "yyyyMMdd_HHmmss");
    const filename = `MVR_Overrides_Export_${timestamp}.csv`;
    
    const file = DriveApp.createFile(filename, csv, MimeType.CSV);
    
    ui.alert('Success', 
      `✅ Overrides Export Complete!\n\n` +
      `File: ${filename}\n` +
      `Overridden tickets: ${filtered.length - 1}\n` +
      `File URL: ${file.getUrl()}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Export failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Export audit log
 */
function exportAuditLog() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const auditSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.OVERRIDE_AUDIT_LOG);
    
    if (!auditSheet) {
      ui.alert('Error', 'Audit Log sheet not found.\n\nNo overrides have been logged yet.', ui.ButtonSet.OK);
      return;
    }
    
    const data = auditSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert('No Data', 'Audit log is empty.', ui.ButtonSet.OK);
      return;
    }
    
    // Convert to CSV
    const csv = data.map(row => 
      row.map(cell => {
        const str = String(cell);
        if (str.includes(',') || str.includes('"') || str.includes('\n')) {
          return '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
      }).join(',')
    ).join('\n');
    
    const timestamp = Utilities.formatDate(new Date(), TIME_CONFIG.TIMEZONE, "yyyyMMdd_HHmmss");
    const filename = `MVR_AuditLog_Export_${timestamp}.csv`;
    
    const file = DriveApp.createFile(filename, csv, MimeType.CSV);
    
    ui.alert('Success', 
      `✅ Audit Log Export Complete!\n\n` +
      `File: ${filename}\n` +
      `Records: ${data.length - 1}\n` +
      `File URL: ${file.getUrl()}`,
      ui.ButtonSet.OK);
    
  } catch (e) {
    ui.alert('Error', `Export failed: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Show documentation
 */
function showDocumentation() {
  const message = `
📚 MVR TICKET TRACKER DOCUMENTATION

SETUP:
1. Configure Freshdesk API credentials
2. (Optional) Setup automatic daily triggers

USAGE:
• Fetch & Append New Tickets - Manual update
• Generate Reports - Create partner summaries
• Full Update - Fetch + Reports in one step

DATA SHEETS:
• MVR_Ticket_History - All historical tickets (42 columns)
• By_Partner_Summary - Partner-level statistics
• Monthly_MVR_Report - Finance billing breakdown
• Executive_Dashboard - High-level KPIs
• Finance_Audit_Trail - Billable tickets for audit
• Ops_Performance - Agent-level metrics
• Override_Audit_Log - Manual override tracking
• Tag_Outcome_Mappings - Dynamic tag classification rules

NEW FEATURES:
• Outcome Classification - Auto-detect MVR results
• Vendor Group Detection - CERTN, PENNDOT, etc.
• Conversation Fetching - Full ticket history
• Manual Overrides - Edit outcomes with audit trail
• Raw Data Export - CSV exports for analysts
• Dynamic Tag Mappings - Configure tag rules without code

SCHEDULE:
Daily automatic update at 8:00 AM CST (if triggers enabled)

FEATURES:
✅ 10-day lookback window
✅ Automatic deduplication
✅ Complete ticket enrichment
✅ Partner grouping reports
✅ Historical append-only record

For detailed documentation, see README.md
  `;
  
  SpreadsheetApp.getUi().alert('Documentation', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════════
// TAG MAPPING MANAGEMENT FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Navigate to Tag Outcome Mappings sheet
 */
function viewTagMappings() {
  // Ensure sheet exists by loading mappings
  loadTagMappings();
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  if (sheet) {
    sheet.activate();
  }
}

/**
 * Add a new tag mapping via dialog
 */
function addTagMapping() {
  const ui = SpreadsheetApp.getUi();
  
  // Get pattern
  const patternResp = ui.prompt('Add Tag Mapping', 'Enter tag pattern (e.g., "suspended" or "dmv[\\s_-]?down"):', ui.ButtonSet.OK_CANCEL);
  if (patternResp.getSelectedButton() !== ui.Button.OK) return;
  const pattern = patternResp.getResponseText().trim();
  if (!pattern) {
    ui.alert('Error', 'Pattern cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  // Get match type
  const matchTypeResp = ui.prompt('Add Tag Mapping', 'Match type: "contains" (simple text) or "regex" (regular expression):', ui.ButtonSet.OK_CANCEL);
  if (matchTypeResp.getSelectedButton() !== ui.Button.OK) return;
  const matchType = matchTypeResp.getResponseText().trim().toLowerCase();
  if (matchType !== 'contains' && matchType !== 'regex') {
    ui.alert('Error', 'Match type must be "contains" or "regex".', ui.ButtonSet.OK);
    return;
  }
  
  // Get outcome type
  const outcomeTypes = Object.values(OUTCOME_TYPES).join(', ');
  const outcomeResp = ui.prompt('Add Tag Mapping', `Outcome type (${outcomeTypes}):`, ui.ButtonSet.OK_CANCEL);
  if (outcomeResp.getSelectedButton() !== ui.Button.OK) return;
  const outcomeType = outcomeResp.getResponseText().trim();
  if (!Object.values(OUTCOME_TYPES).includes(outcomeType)) {
    ui.alert('Error', `Invalid outcome type. Must be one of: ${outcomeTypes}`, ui.ButtonSet.OK);
    return;
  }
  
  // Get priority
  const priorityResp = ui.prompt('Add Tag Mapping', 'Priority (1-99, lower = higher priority):', ui.ButtonSet.OK_CANCEL);
  if (priorityResp.getSelectedButton() !== ui.Button.OK) return;
  const priority = parseInt(priorityResp.getResponseText().trim()) || 50;
  
  // Get billable
  const billableResp = ui.alert('Add Tag Mapping', 'Is this outcome billable?', ui.ButtonSet.YES_NO);
  const isBillable = billableResp === ui.Button.YES;
  
  // Get notes
  const notesResp = ui.prompt('Add Tag Mapping', 'Notes (optional):', ui.ButtonSet.OK_CANCEL);
  const notes = notesResp.getSelectedButton() === ui.Button.OK ? notesResp.getResponseText().trim() : '';
  
  // Add to sheet - follows TAG_MAPPING_HEADERS order from Config.gs
  loadTagMappings(); // Ensure sheet exists
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  sheet.appendRow([pattern, matchType, outcomeType, priority, isBillable, true, notes]);
  
  ui.alert('Success', `✅ Tag mapping added:\n\nPattern: ${pattern}\nOutcome: ${outcomeType}\nBillable: ${isBillable}`, ui.ButtonSet.OK);
  sheet.activate();
}

/**
 * Test a tag string against current mappings
 */
function testTagMatch() {
  const ui = SpreadsheetApp.getUi();
  
  const tagResp = ui.prompt('Test Tag Match', 'Enter tags to test (comma-separated, e.g., "suspended, reviewed"):', ui.ButtonSet.OK_CANCEL);
  if (tagResp.getSelectedButton() !== ui.Button.OK) return;
  
  const tagString = tagResp.getResponseText().trim();
  if (!tagString) {
    ui.alert('Error', 'Please enter tags to test.', ui.ButtonSet.OK);
    return;
  }
  
  // Convert to array and classify
  const tags = tagString.split(',').map(t => t.trim());
  const result = classifyOutcome(tags, []);
  
  ui.alert('Test Result', 
    `Tags: ${tags.join(', ')}\n\n` +
    `Outcome: ${result.outcome}\n` +
    `Source: ${result.source}\n` +
    `Billable: ${result.isBillable}\n` +
    `Notes: ${result.notes}`,
    ui.ButtonSet.OK);
}

/**
 * Reset tag mappings to defaults
 */
function resetTagMappingsToDefaults() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert('Reset Tag Mappings', 
    '⚠️ This will DELETE all current tag mappings and restore defaults.\n\nAre you sure?', 
    ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // Reload will recreate with defaults
  loadTagMappings();
  
  ui.alert('Success', `✅ Tag mappings reset to ${DEFAULT_TAG_MAPPINGS.length} defaults.`, ui.ButtonSet.OK);
  
  ss.getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS).activate();
}

/**
 * Load recommended tag mappings - NOW CALLS syncTagMappings()
 * This is a wrapper for backward compatibility with menu
 * The actual logic is in syncTagMappings() in OutcomeTracking.gs
 */
function loadRecommendedMappings() {
  // Delegate to syncTagMappings() which is the single source of truth
  syncTagMappings();
}

/**
 * Apply data validation dropdown for Outcome Override column
 */
function applyOutcomeValidation() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  if (!sheet) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const overrideColIdx = headers.indexOf('Outcome Override') + 1;
  
  if (overrideColIdx === 0) {
    Logger.log('Outcome Override column not found');
    return;
  }
  
  // Get valid outcomes from OUTCOME_TYPES
  const outcomeValues = Object.values(OUTCOME_TYPES);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(outcomeValues, true)
    .setAllowInvalid(false)
    .build();
  
  // Apply to all data rows
  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(2, overrideColIdx, lastRow - 1, 1).setDataValidation(rule);
  
  Logger.log(`✅ Applied outcome validation to column ${overrideColIdx}`);
}

/**
 * View the Discovered Tags sheet
 */
function viewDiscoveredTags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.DISCOVERED_TAGS);
  
  if (!sheet) {
    // First time - run discovery to create sheet
    SpreadsheetApp.getActiveSpreadsheet().toast('Running tag discovery...', 'Please wait', 3);
    discoverTagsFromRaw();
    sheet = ss.getSheetByName(SHEET_NAMES.DISCOVERED_TAGS);
  }
  
  if (sheet) {
    sheet.activate();
  } else {
    SpreadsheetApp.getUi().alert('Error', 'Could not find or create Discovered Tags sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Show dialog with unmapped tags and mapping options
 */
function viewUnmappedTags() {
  const unmappedTags = identifyUnmappedTags();
  
  if (unmappedTags.length === 0) {
    SpreadsheetApp.getUi().alert('All Clear!', '✅ All discovered tags have outcome mappings.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Build HTML dialog
  let html = '<div style="font-family: Arial, sans-serif; padding: 10px;">';
  html += '<h2>🏷️ Unmapped Tags</h2>';
  html += '<p>These tags were found in raw tickets but have no outcome mapping:</p>';
  html += '<table style="width: 100%; border-collapse: collapse; margin-top: 10px;">';
  html += '<tr style="background: #f0f0f0; font-weight: bold;">';
  html += '<th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Tag</th>';
  html += '<th style="padding: 8px; border: 1px solid #ddd; text-align: right;">Usage Count</th>';
  html += '<th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Suggested Outcome</th>';
  html += '</tr>';
  
  unmappedTags.slice(0, 20).forEach(tag => {
    const suggestion = suggestOutcomeForTag(tag.tag);
    html += `<tr>`;
    html += `<td style="padding: 8px; border: 1px solid #ddd;">${tag.tag}</td>`;
    html += `<td style="padding: 8px; border: 1px solid #ddd; text-align: right;">${tag.count}</td>`;
    html += `<td style="padding: 8px; border: 1px solid #ddd;">${suggestion || '(none)'}</td>`;
    html += `</tr>`;
  });
  
  html += '</table>';
  html += '<p style="margin-top: 15px;">Use <b>🏷️ Tag Management → Map New Tag</b> to create mappings.</p>';
  html += '<button onclick="google.script.host.close()" style="margin-top: 10px; padding: 8px 16px; background: #4285f4; color: white; border: none; border-radius: 4px; cursor: pointer;">Close</button>';
  html += '</div>';
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Unmapped Tags');
}

/**
 * Wizard to map a new tag to an outcome
 */
function mapNewTag() {
  const ui = SpreadsheetApp.getUi();
  
  // Step 1: Get tag pattern
  const tagResponse = ui.prompt(
    'Map New Tag - Step 1/4',
    'Enter the tag pattern to map (e.g., "violation" or "*suspend*"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (tagResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const tagPattern = tagResponse.getResponseText().trim();
  if (!tagPattern) {
    ui.alert('Error', 'Tag pattern cannot be empty.', ui.ButtonSet.OK);
    return;
  }
  
  // Step 2: Get suggested outcome
  const suggestion = suggestOutcomeForTag(tagPattern);
  const outcomeHint = suggestion ? `\n\nSuggested: ${suggestion}` : '';
  const outcomeList = Object.values(OUTCOME_TYPES).join(', ');
  
  const outcomeResponse = ui.prompt(
    'Map New Tag - Step 2/4',
    `Select outcome for "${tagPattern}":${outcomeHint}\n\nValid outcomes: ${outcomeList}`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (outcomeResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const outcome = outcomeResponse.getResponseText().trim();
  if (!Object.values(OUTCOME_TYPES).includes(outcome)) {
    ui.alert('Error', `Invalid outcome. Must be one of: ${outcomeList}`, ui.ButtonSet.OK);
    return;
  }
  
  // Step 3: Billable?
  const billableResponse = ui.alert(
    'Map New Tag - Step 3/4',
    `Is "${tagPattern}" billable?`,
    ui.ButtonSet.YES_NO
  );
  
  const isBillable = (billableResponse === ui.Button.YES);
  
  // Step 4: Priority (optional)
  const priorityResponse = ui.prompt(
    'Map New Tag - Step 4/4',
    'Enter priority (1-100, higher = checked first). Leave blank for default 50:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (priorityResponse.getSelectedButton() !== ui.Button.OK) return;
  
  let priority = 50;
  const priorityText = priorityResponse.getResponseText().trim();
  if (priorityText) {
    priority = parseInt(priorityText, 10);
    if (isNaN(priority) || priority < 1 || priority > 100) {
      ui.alert('Error', 'Priority must be a number between 1 and 100.', ui.ButtonSet.OK);
      return;
    }
  }
  
  // Check for duplicate pattern
  const existingMappings = loadTagMappings();
  const duplicate = existingMappings.find(m => m.pattern.toLowerCase() === tagPattern.toLowerCase());
  
  if (duplicate) {
    const overwrite = ui.alert(
      'Duplicate Pattern',
      `Pattern "${tagPattern}" already exists with outcome "${duplicate.outcomeType}".

Do you want to overwrite it?`,
      ui.ButtonSet.YES_NO
    );
    
    if (overwrite !== ui.Button.YES) {
      return;
    }
  }
  
  // Add mapping
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.TAG_OUTCOME_MAPPINGS);
  
  if (!sheet) {
    ui.alert('Error', 'Tag Outcome Mappings sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  // If overwriting, remove old mapping first
  if (duplicate) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toLowerCase() === tagPattern.toLowerCase()) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
  }
  
  // Append new mapping
  sheet.appendRow([
    tagPattern,
    outcome,
    isBillable ? 'Yes' : 'No',
    priority,
    new Date(),
    `Mapped via wizard`
  ]);
  
  // Clear cache to reload mappings
  CacheService.getScriptCache().remove('TAG_MAPPINGS_CACHE');
  
  ui.alert(
    'Success',
    `✅ Tag mapping added:\n\n` +
    `Pattern: ${tagPattern}\n` +
    `Outcome: ${outcome}\n` +
    `Billable: ${isBillable ? 'Yes' : 'No'}\n` +
    `Priority: ${priority}`,
    ui.ButtonSet.OK
  );
  
  sheet.activate();
}

/**
 * Reapply tag mappings to all history tickets
 * Reclassifies tickets without manual overrides using current mappings
 */
function reapplyTagMappingsToHistory() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    'Reapply Tag Mappings',
    '⚠️ This will reclassify ALL tickets in MVR_Ticket_History using current tag mappings.\n\n' +
    'Tickets with manual Outcome Overrides will NOT be changed.\n\n' +
    'This may take several minutes. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Reapplying tag mappings...', 'Please wait', -1);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.MVR_TICKET_HISTORY);
  
  if (!sheet) {
    ui.alert('Error', 'MVR_Ticket_History sheet not found.', ui.ButtonSet.OK);
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const tagsColIdx = headers.indexOf('Tags') + 1;
  const outcomeColIdx = headers.indexOf('MVR Outcome') + 1;
  const overrideColIdx = headers.indexOf('Outcome Override') + 1;
  
  if (tagsColIdx === 0 || outcomeColIdx === 0) {
    ui.alert('Error', 'Required columns not found in history sheet.', ui.ButtonSet.OK);
    return;
  }
  
  // Get all data
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Info', 'No tickets to reclassify.', ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const data = dataRange.getValues();
  
  let updatedCount = 0;
  
  // Process each row
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const hasOverride = overrideColIdx > 0 && row[overrideColIdx - 1];
    
    // Skip if manual override exists
    if (hasOverride) continue;
    
    const tags = row[tagsColIdx - 1] || '';
    const tagArray = tags.split(',').map(t => t.trim()).filter(t => t);
    
    // Classify using current mappings
    const result = classifyOutcome(tagArray, []);
    
    // Update outcome if different
    const currentOutcome = row[outcomeColIdx - 1];
    
    if (currentOutcome !== result.outcome) {
      row[outcomeColIdx - 1] = result.outcome;
      updatedCount++;
    }
    
    // Progress indicator every 100 tickets
    if ((i + 1) % 100 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Processing: ${i + 1}/${data.length}`, 'Reapplying...', 1);
    }
  }
  
  // Write back all data
  if (updatedCount > 0) {
    dataRange.setValues(data);
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Complete!', 'Tag Mappings Reapplied', 3);
  
  ui.alert(
    'Complete',
    `✅ Reapplied tag mappings to history.\n\n` +
    `Tickets processed: ${data.length}\n` +
    `Tickets updated: ${updatedCount}\n` +
    `Tickets skipped (manual override): ${data.length - updatedCount}`,
    ui.ButtonSet.OK
  );
}


