/**
 * 📊 System Health Monitor
 * Used by Jules CLI and external monitoring tools
 */
function runHealthCheck() {
    const results = {
        timestamp: new Date(),
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
        const response = UrlFetchApp.fetch('https://www.google.com');
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
        const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET.ID);
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

    // Log results
    console.log(JSON.stringify(results, null, 2));
    return results;
}
