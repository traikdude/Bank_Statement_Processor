/**
 * 📊 System Health Monitor
 * Used by Jules CLI and external monitoring tools
 * 
 * NOTE: This file is kept for backward compatibility.
 * The primary runHealthCheck() is now in Code.js with full CONFIG access.
 * This version uses hardcoded IDs as a standalone fallback.
 */

// Hardcoded for standalone execution (CONFIG is in Code.js)
const MONITOR_SPREADSHEET_ID = '1XuvPyWNhB3WAOMHDO9wXkZ5AO36Cce13PH8PAojG9Eo';
const MONITOR_INPUT_FOLDER_ID = '1vVQC5F8ZrKZnqv5QAWcye5VKGxCQ1c3q';

/**
 * Standalone health check function (fallback version)
 * @returns {Object} Health check results
 */
function runMonitorHealthCheck() {
    const results = {
        timestamp: new Date(),
        source: 'monitoring.js',
        tests: []
    };

    // Test 1: Script Execution
    try {
        results.tests.push({
            name: 'Script Execution',
            status: 'PASS',
            message: 'Script executes successfully'
        });
    } catch (error) {
        results.tests.push({
            name: 'Script Execution',
            status: 'FAIL',
            message: error.toString()
        });
    }

    // Test 2: API Connectivity
    try {
        const response = UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
        results.tests.push({
            name: 'API Connectivity',
            status: response.getResponseCode() === 200 ? 'PASS' : 'FAIL',
            message: `HTTP ${response.getResponseCode()}`
        });
    } catch (error) {
        results.tests.push({
            name: 'API Connectivity',
            status: 'FAIL',
            message: error.toString()
        });
    }

    // Test 3: Spreadsheet Connectivity
    try {
        const ss = SpreadsheetApp.openById(MONITOR_SPREADSHEET_ID);
        results.tests.push({
            name: 'Spreadsheet Access',
            status: 'PASS',
            message: `Connected to ${ss.getName()}`
        });
    } catch (error) {
        results.tests.push({
            name: 'Spreadsheet Access',
            status: 'FAIL',
            message: error.toString()
        });
    }

    // Overall status
    results.overall = results.tests.every(t => t.status === 'PASS') ? 'HEALTHY' : 'DEGRADED';

    // Log results
    console.log('📊 Monitor Health Check:', JSON.stringify(results, null, 2));
    return results;
}
