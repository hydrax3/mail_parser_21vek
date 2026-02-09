const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
// Helper to determine if app is running from source or packaged
const isPackaged = process.mainModule && process.mainModule.filename.indexOf('app.asar') !== -1 || (process.resourcesPath && fs.existsSync(path.join(process.resourcesPath, 'app.asar')));

// Determine base path for configs
const basePath = isPackaged
    ? process.resourcesPath
    : path.join(__dirname, '..', '..');

require('dotenv').config({ path: path.join(basePath, '.env') });

// Configuration
const CREDENTIALS_PATH = path.join(basePath, 'credentials.json');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];

let sheetsClient = null;

/**
 * Initialize Google Sheets API client
 */
async function initClient() {
    if (sheetsClient) return sheetsClient;

    try {
        // Check if credentials file exists
        if (!fs.existsSync(CREDENTIALS_PATH)) {
            throw new Error(`credentials.json not found at ${CREDENTIALS_PATH}`);
        }

        const auth = new google.auth.GoogleAuth({
            keyFile: CREDENTIALS_PATH,
            scopes: SCOPES
        });

        sheetsClient = google.sheets({ version: 'v4', auth });
        return sheetsClient;
    } catch (error) {
        console.error('Failed to initialize Google Sheets client:', error);
        throw error;
    }
}

/**
 * Extract spreadsheet ID from URL
 */
function extractSpreadsheetId(url) {
    if (!url) return null;
    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : null;
}

/**
 * Fetch all emails from Google Sheet
 * Returns array of objects with: id, subject, sender, time, status, type, lastReplyer
 */
async function getEmails() {
    try {
        const client = await initClient();
        const sheetUrl = process.env.GOOGLE_SHEET_URL;

        if (!sheetUrl) {
            throw new Error('GOOGLE_SHEET_URL not found in .env');
        }

        const spreadsheetId = extractSpreadsheetId(sheetUrl);
        if (!spreadsheetId) {
            throw new Error('Invalid Google Sheet URL');
        }

        // Get sheet name for GID=0 (main emails sheet)
        const metadata = await client.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });
        const mainSheet = metadata.data.sheets.find(s => s.properties.sheetId === 0);
        const sheetName = mainSheet ? mainSheet.properties.title : 'Sheet1';

        const response = await client.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:H` // id, theme_of_mail, sender, time, status_of_reply, type_of_email, last_replyer, reminder_status
        });

        const rows = response.data.values || [];
        if (rows.length <= 1) {
            return []; // Only header or empty
        }

        // Skip header row
        const emails = rows.slice(1).map((row, index) => ({
            id: row[0] || '',
            subject: row[1] || '',
            sender: row[2] || '',
            time: row[3] || '',
            status: row[4] || '',
            type: row[5] || '',
            lastReplyer: row[6] || '',
            reminderStatus: row[7] || '',
            rowIndex: index + 2 // For reference (1-indexed, after header)
        }));

        return emails;
    } catch (error) {
        console.error('Error fetching emails:', error);
        throw error;
    }
}

/**
 * Check if an email requires a notification based on criteria:
 * 1. Status is 'ответа нет' or 'ответ не от оператора'.
 * 2. Time elapsed > 6 hours.
 * 3. Status is NOT 'закрыт'.
 * Returns an object { notify: boolean, reason: string }
 */
function checkNotificationCriteria(email) {
    // 1. Check if manually closed
    if (email.reminderStatus && email.reminderStatus.toLowerCase().trim() === 'закрыт') {
        return { notify: false, reason: 'closed' };
    }

    // 2. Check status
    const status = email.status ? email.status.toLowerCase().trim() : '';
    const relevantStatuses = ['ответа нет', 'ответ не от оператора'];

    if (!relevantStatuses.includes(status)) {
        return { notify: false, reason: 'status_ok' };
    }

    // 3. Determine time reference
    // Use reminderStatus (Col H) if it's a date, otherwise fallback to email.time (Col D)
    let timeStr = email.reminderStatus;

    // If Col H is empty or not a date (and not 'закрыт'), fall back to Col D
    // Note: parser.py fills Col H on creation, so it should be there.
    if (!timeStr || timeStr.trim() === '') {
        timeStr = email.time;
    }

    try {
        const refDate = new Date(timeStr);
        if (isNaN(refDate.getTime())) {
            // Try fallback if H was invalid
            const fallbackDate = new Date(email.time);
            if (!isNaN(fallbackDate.getTime())) {
                // Use D logic
                // But wait, if H is garbage, we might want to ignore or fallback.
                // Let's assume fallback to D is safer.
                if (checkTimeDiff(fallbackDate, 6)) {
                    return { notify: true, reason: status };
                }
            }
            return { notify: false, reason: 'invalid_date' };
        }

        if (checkTimeDiff(refDate, 6)) {
            return { notify: true, reason: status };
        }
    } catch (e) {
        return { notify: false, reason: 'error_parsing_date' };
    }

    return { notify: false, reason: 'under_threshold' };
}

function checkTimeDiff(dateObj, hours) {
    const now = new Date();
    const diffMs = now - dateObj;
    const diffHours = diffMs / (1000 * 60 * 60);
    return diffHours > hours;
}

/**
 * Calculate if email is overdue (no reply for more than 24 hours) - LEGACY / UPDATED
 * Keeps 24h for legacy or redirects to new logic? 
 * The task implies we want 6h logic. I'll make isOverdue use the new logic but strictly for boolean.
 */
function isOverdue(email) {
    const result = checkNotificationCriteria(email);
    return result.notify;
}

/**
 * Get overdue emails
 */
async function getOverdueEmails() {
    const emails = await getEmails();
    return emails.filter(isOverdue);
}

/**
 * Get list of operators from the Operators sheet (GID=2012399964)
 */
async function getOperators() {
    try {
        const client = await initClient();
        const sheetUrl = process.env.GOOGLE_SHEET_URL;
        const spreadsheetId = extractSpreadsheetId(sheetUrl);

        // First, get sheet metadata to find the sheet name by GID
        const metadata = await client.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });

        const operatorsSheet = metadata.data.sheets.find(
            s => s.properties.sheetId === 2115150025
        );

        if (!operatorsSheet) {
            throw new Error('Operators sheet (GID=2115150025) not found');
        }

        const sheetName = operatorsSheet.properties.title;

        const response = await client.spreadsheets.values.get({
            spreadsheetId,
            range: `'${sheetName}'!A:A`
        });

        return (response.data.values || []).flat().filter(e => e && e.trim());
    } catch (error) {
        console.error('Error fetching operators:', error);
        throw error;
    }
}

/**
 * Update reminder status for an email row
 */
async function updateReminderStatus(rowIndex, status) {
    try {
        const client = await initClient();
        const sheetUrl = process.env.GOOGLE_SHEET_URL;
        const spreadsheetId = extractSpreadsheetId(sheetUrl);

        // Get sheet name for GID=0 (main emails sheet)
        const metadata = await client.spreadsheets.get({
            spreadsheetId,
            fields: 'sheets.properties'
        });
        const mainSheet = metadata.data.sheets.find(s => s.properties.sheetId === 0);
        const sheetName = mainSheet ? mainSheet.properties.title : 'Sheet1';

        await client.spreadsheets.values.update({
            spreadsheetId,
            range: `'${sheetName}'!H${rowIndex}`,
            valueInputOption: 'RAW',
            requestBody: { values: [[status]] }
        });

        return { success: true };
    } catch (error) {
        console.error('Error updating reminder status:', error);
        throw error;
    }
}

/**
 * Test function for verification
 */
async function test() {
    try {
        console.log('Testing Google Sheets connection...');
        const emails = await getEmails();
        console.log(`✅ Found ${emails.length} emails`);

        const overdue = emails.filter(isOverdue);
        console.log(`⚠️ ${overdue.length} overdue emails (>6h)`);

        if (emails.length > 0) {
            console.log('\nSample email:');
            console.log(JSON.stringify(emails[0], null, 2));
            console.log('Notification Check:', checkNotificationCriteria(emails[0]));
        }

        return { success: true, count: emails.length, overdue: overdue.length };
    } catch (error) {
        console.error('❌ Test failed:', error.message);
        return { success: false, error: error.message };
    }
}

module.exports = {
    initClient,
    getEmails,
    getOverdueEmails,
    isOverdue,
    checkNotificationCriteria, // Exported
    getOperators,
    updateReminderStatus,
    test
};
