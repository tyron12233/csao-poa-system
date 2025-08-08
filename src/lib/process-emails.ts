"use client";

interface ActivityDetails {
    title: string;
    organization: string;
    startDate: string;
    endDate: string;
    description: string;
    time: string;
    venue: string;
    type: string;
}

// --- Types ---
export interface UserProfile { email: string; name: string; picture: string; }
export type EmailTaskStatus = 'queued' | 'fetching' | 'parsing' | 'building_request' | 'writing' | 'done' | 'skipped' | 'error';
export interface EmailTask { id: string; subject: string; status: EmailTaskStatus; error?: string; }
export type ParsedSuccessResult = { status: 'success', data: { messageId: string, subject: string, parsedData: any, monthSheetNames: string[], rowData: any[] } };
export type ParsedResult = ParsedSuccessResult | { status: 'error', messageId: string, error: string };

// --- Config ---
const LOG_SHEET_NAME = 'POA Log';
const MONTHLY_HEADERS = ['Categories', 'Name Of Organization', 'Title Of Activity', 'Description', 'Start Date Of Implementation', 'End Date Of Implementation', 'Time', 'Venue', 'Type Of Activity', 'Link Of Approved POA', 'Narrative Report'];
const LOG_HEADERS = ['Timestamp', 'Message ID', 'Status', 'Organization', 'Title'];
const CONDITIONAL_FORMAT_RULES: Record<string, { bg: any; fg: any }> = {
    UNSET: { bg: { red: 0.878, green: 0.878, blue: 0.878 }, fg: { red: 0, green: 0, blue: 0 } },
    SPIN: { bg: { red: 0.733, green: 0.871, blue: 0.984 }, fg: { red: 0.051, green: 0.278, blue: 0.631 } },
    SCRO: { bg: { red: 1.0, green: 0.878, blue: 0.698 }, fg: { red: 0.902, green: 0.318, blue: 0 } },
    PROF: { bg: { red: 0.784, green: 0.902, blue: 0.788 }, fg: { red: 0.106, green: 0.369, blue: 0.125 } },
};

// --- Pure helpers ---
const decodeBase64Url = (encoded: string): string => {
    try {
        return decodeURIComponent(
            atob(encoded.replace(/-/g, '+').replace(/_/g, '/'))
                .split('')
                .map(c => '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2))
                .join('')
        );
    } catch {
        return '';
    }
};

const getEmailBodyHtml = (payload: any): string => {
    function find(parts: any[]): string | null {
        for (const p of parts) {
            if (p.mimeType === 'text/html' && p.body?.data) return decodeBase64Url(p.body.data);
            if (p.parts) { const n = find(p.parts); if (n) return n; }
        }
        return null;
    }
    if (payload.mimeType === 'text/html' && payload.body?.data) return decodeBase64Url(payload.body.data);
    if (payload.parts) return find(payload.parts) || '';
    return '';
};

const extractDataFromHtml = (html: string, label: string): string => {
    const m = html.match(new RegExp(`>${label}</td>\\s*<td[^>]*>(.*?)</td>`, 's'));
    if (m && m[1]) {
        const el = document.createElement('div');
        el.innerHTML = m[1];
        return el.textContent?.trim() || 'Not Found';
    }
    return 'Not Found';
};

/**
 * Parses event details from an HTML email body using DOM parsing with TypeScript.
 *
 * @param {string} htmlBody The HTML content of the email.
 * @returns {Partial<ActivityDetails>} An object containing the extracted event details.
 * It's a 'Partial' type because some fields may not be found during parsing.
 */
function parseActivityDetailsFromHtml(htmlBody: string): Partial<ActivityDetails> {
    // This map links HTML labels to keys of our ActivityDetails interface.
    // Using `keyof ActivityDetails` ensures we can only map to valid property names.
    const keyMap: { [key: string]: keyof ActivityDetails } = {
        'Title of Activity:': 'title',
        'Name of Organization:': 'organization',
        'Start Date of Implementation:': 'startDate',
        'End Date of Implementation:': 'endDate',
        'Rationale/Brief Description:': 'description',
        'Time of Implementation:': 'time',
        'Venue/Platform:': 'venue',
        'Type of Implementation (Online -Social Media Posting; Google Meet; Zoom etc)/Face to Face:': 'type',
    };

    // The details object is typed as a Partial, meaning it will have some or all
    // of the properties from ActivityDetails.
    const details: Partial<ActivityDetails> = {};

    try {
        const doc = Document.parseHTMLUnsafe(htmlBody);

        // Select all rows from the table containing the activity details.
        const rows: Element[] = Array.from(doc.querySelectorAll('.response-items tr'));

        rows.forEach(row => {
            const cells = Array.from(row.querySelectorAll('td'));

            if (cells.length >= 2) {
                const label = (cells[0].textContent || '').trim();

                // Check if the extracted label is a key in our map.
                if (label in keyMap) {
                    const value = (cells[1].textContent || '').trim();
                    const objectKey = keyMap[label];
                    details[objectKey] = value;
                }
            }
        });
    } catch (e: unknown) {
        // It's best practice to type the catch variable as 'unknown'.
        const errorMessage = e instanceof Error ? e.message : String(e);
        console.error(`Failed to parse HTML: ${errorMessage}`);
        // Return whatever details were successfully parsed before the error.
        return details;
    }

    return details;
}



const parseEmailBody = (htmlBody: string) => {
    const details = parseActivityDetailsFromHtml(htmlBody);
    return {
        title: details.title || 'Not Found',
        organization: details.organization || 'Not Found',
        startDate: details.startDate || 'Not Found',
        endDate: details.endDate || 'Not Found',
        description: details.description || 'Not Found',
        time: details.time || 'Not Found',
        venue: details.venue || 'Not Found',
        type: details.type || 'Not Found',
    };
};

const parseDateWithInferredYear = (dateString: string): Date | null => {
    if (!dateString || dateString === 'Not Found') {
        return null;
    }
    const d = new Date(dateString);
    if (isNaN(d.getTime())) {
        return null;
    }
    const now = new Date();
    if (d.getFullYear() < now.getFullYear() - 5) {
        let y = now.getFullYear();
        if (d.getMonth() < now.getMonth()) { y += 1; }
        d.setFullYear(y);
    }
    return d;
};

const getMonthsBetweenDates = (start: Date, end: Date): Date[] => {
    const m: Date[] = [];
    let c = new Date(start.getFullYear(), start.getMonth(), 1);
    while (c <= end) { m.push(new Date(c)); c.setMonth(c.getMonth() + 1); }
    return m;
};

const formatMonthYear = (date: Date): string =>
    date.toLocaleString('en-US', { month: 'long', year: 'numeric' }).replace(' ', ' - ');

// --- Sheets helpers ---
const generateRowRequests = (sheetId: number, rowIndex: number, rowData: any[]) => {
    const requests: any[] = [
        { updateCells: { start: { sheetId, rowIndex, columnIndex: 0 }, rows: [{ values: rowData.map(val => ({ userEnteredValue: { stringValue: String(val) } })) }], fields: 'userEnteredValue' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'ROWS', startIndex: rowIndex, endIndex: rowIndex + 1 }, properties: { pixelSize: 160 }, fields: 'pixelSize' } },
        { repeatCell: { range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1 }, cell: { userEnteredFormat: { horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'CLIP' } }, fields: 'userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy)' } },
        { repeatCell: { range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 3, endColumnIndex: 4 }, cell: { userEnteredFormat: { horizontalAlignment: 'LEFT', wrapStrategy: 'WRAP' } }, fields: 'userEnteredFormat(horizontalAlignment,wrapStrategy)' } },
        { setDataValidation: { range: { sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: 1 }, rule: { condition: { type: 'ONE_OF_LIST', values: ['UNSET', 'SPIN', 'SCRO', 'PROF'].map(v => ({ userEnteredValue: v })) }, showCustomUi: true, strict: true } } }
    ];
    Object.entries(CONDITIONAL_FORMAT_RULES).forEach(([text, format]) => {
        requests.push({ addConditionalFormatRule: { rule: { ranges: [{ sheetId, startRowIndex: rowIndex, endRowIndex: rowIndex + 1, startColumnIndex: 0, endColumnIndex: 1 }], booleanRule: { condition: { type: 'TEXT_EQ', values: [{ userEnteredValue: text }] }, format: { textFormat: { foregroundColorStyle: { rgbColor: format.fg } }, backgroundColorStyle: { rgbColor: format.bg } } } }, index: 0 } });
    });
    return requests;
};

const generateHeaderRequests = (sheetId: number, headers: string[]) => {
    const styleRequest = {
        repeatCell: {
            range: { sheetId, startRowIndex: 0, endRowIndex: 1 },
            cell: {
                userEnteredFormat: {
                    backgroundColorStyle: { rgbColor: { red: 0.259, green: 0.522, blue: 0.957 } },
                    textFormat: { foregroundColorStyle: { rgbColor: { red: 1.0, green: 1.0, blue: 1.0 } }, bold: true },
                    horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE',
                }
            },
            fields: 'userEnteredFormat(backgroundColorStyle,textFormat,horizontalAlignment,verticalAlignment)'
        }
    };

    const freezeRowRequest = {
        updateSheetProperties: {
            properties: { sheetId, gridProperties: { frozenRowCount: 1 } },
            fields: 'gridProperties.frozenRowCount'
        }
    };

    const setHeaderDataRequest = {
        updateCells: {
            start: { sheetId, rowIndex: 0, columnIndex: 0 },
            rows: [{ values: headers.map(val => ({ userEnteredValue: { stringValue: val } })) }],
            fields: 'userEnteredValue'
        }
    };

    const columnWidthRequests = headers === MONTHLY_HEADERS ? [
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 0, endIndex: 1 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 1, endIndex: 2 }, properties: { pixelSize: 250 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 2, endIndex: 3 }, properties: { pixelSize: 350 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 3, endIndex: 4 }, properties: { pixelSize: 500 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 4, endIndex: 6 }, properties: { pixelSize: 180 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 6, endIndex: 8 }, properties: { pixelSize: 230 }, fields: 'pixelSize' } },
        { updateDimensionProperties: { range: { sheetId, dimension: 'COLUMNS', startIndex: 8, endIndex: 11 }, properties: { pixelSize: 250 }, fields: 'pixelSize' } },
    ] : [];

    return [setHeaderDataRequest, styleRequest, freezeRowRequest, ...columnWidthRequests];
};

// --- Email parsing ---
async function fetchAndParseEmail(gapi: any, messageId: string, callbacks: Callbacks): Promise<ParsedResult> {
    try {
        callbacks.updateTask(messageId, { status: 'fetching' });
        const { result } = await gapi.client.gmail.users.messages.get({ userId: 'me', id: messageId });
        const subject = result.payload.headers.find((h: any) => h.name.toLowerCase() === 'subject')?.value || 'No Subject';
        callbacks.updateTask(messageId, { subject, status: 'parsing' });

        const htmlBody = getEmailBodyHtml(result.payload);
        if (!htmlBody) throw new Error('HTML body not found.');

        const parsedData = parseEmailBody(htmlBody);
        const startDate = parseDateWithInferredYear(parsedData.startDate);
        const endDate = parseDateWithInferredYear(parsedData.endDate);
        if (!startDate || !endDate) throw new Error(`Invalid dates for: ${parsedData.title}`);

        const monthSheetNames = getMonthsBetweenDates(startDate, endDate).map(formatMonthYear);
        const rowData = ['UNSET', parsedData.organization, parsedData.title, parsedData.description, parsedData.startDate, parsedData.endDate, parsedData.time, parsedData.venue, parsedData.type, '', ''];

        return { status: 'success', data: { messageId, subject, parsedData, monthSheetNames, rowData } };
    } catch (err: any) {
        const errorMessage = err.result?.error?.message || err.message || 'Unknown parsing error.';
        callbacks.updateTask(messageId, { status: 'error', error: errorMessage });
        return { status: 'error', messageId, error: errorMessage };
    }
}

// --- Callback contracts ---
export type Callbacks = {
    status: (msg: string) => void;
    queue: (tasks: EmailTask[]) => void;
    updateTask: (id: string, patch: Partial<EmailTask>) => void;
    updateTasksBulk: (ids: string[], patch: Partial<EmailTask>) => void;
};

export type ProcessEmailsParams = {
    spreadsheetId: string;
    senderEmail: string;
    useDateFilter: boolean;
    callbacks: Callbacks;
};

export async function processEmails(gapi: any, params: ProcessEmailsParams): Promise<{ processed: number }> {
    const { spreadsheetId, senderEmail, useDateFilter, callbacks } = params;

    // --- Phase 1: Fetch emails ---
    let gmailQuery = `from:${senderEmail} "POA" "The request is now complete."`;
    const processedIds = new Set<string>();

    if (useDateFilter) {
        try {
            const logData = await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId, range: `${LOG_SHEET_NAME}!A2:B` });
            const logRows: any[] = logData.result.values || [];
            if (logRows.length > 0) {
                const latestTimestamp = logRows.reduce((latest: Date | null, row: any[]) => {
                    if (row && row[0]) {
                        const current = new Date(row[0]);
                        if (!isNaN(current.getTime()) && (!latest || current > latest)) return current;
                    }
                    return latest;
                }, null as Date | null);
                if (latestTimestamp) {
                    const unixTimestamp = Math.floor(latestTimestamp.getTime() / 1000);
                    gmailQuery += ` after:${unixTimestamp}`;
                    callbacks.status(`Optimized: Fetching only emails after ${latestTimestamp.toLocaleString()}...`);
                }
            }
            logRows.forEach(row => { if (row && row[1]) processedIds.add(row[1]); });
        } catch (e: any) {
            if (e.result?.error?.code === 400) {
                callbacks.status('POA Log sheet not found. Fetching all emails...');
            } else {
                throw e;
            }
        }
    }

    const listRes = await gapi.client.gmail.users.messages.list({ userId: 'me', q: gmailQuery });
    const newMessages = (listRes.result.messages || []).filter((msg: any) => !processedIds.has(msg.id!));
    if (newMessages.length === 0) {
        callbacks.status('No new emails to process.');
        return { processed: 0 };
    }

    // --- Phase 2: Parse emails ---
    callbacks.queue(newMessages.map((msg: any) => ({ id: msg.id!, subject: 'In queue...', status: 'queued' })));
    const parsedResults: ParsedResult[] = [];
    for (let i = 0; i < newMessages.length; i += 5) {
        const chunk = newMessages.slice(i, i + 5);
        parsedResults.push(...await Promise.all(chunk.map((msg: any) => fetchAndParseEmail(gapi, msg.id!, callbacks))));
    }
    const successfulResults = parsedResults.filter((p): p is ParsedSuccessResult => p.status === 'success');
    const successfulIds = successfulResults.map(p => p.data.messageId);
    callbacks.updateTasksBulk(successfulIds, { status: 'building_request' });

    // --- Phase 3: Pre-compute sheet state ---
    callbacks.status('Pre-computing sheet states...');
    const dataBySheet = new Map<string, any[][]>();
    const neededSheetNames = new Set<string>([LOG_SHEET_NAME]);
    for (const result of successfulResults) {
        for (const sheetName of result.data.monthSheetNames) {
            if (!dataBySheet.has(sheetName)) dataBySheet.set(sheetName, []);
            dataBySheet.get(sheetName)!.push(result.data.rowData);
            neededSheetNames.add(sheetName);
        }
    }

    const initialSheetProps = await gapi.client.sheets.spreadsheets.get({ spreadsheetId });
    const existingSheets = new Map<string, any>(initialSheetProps.result.sheets?.map((s: any) => [s.properties.title, { sheetId: s.properties.sheetId }]));
    const sheetsToCreate = [...neededSheetNames].filter(name => !existingSheets.has(name));

    if (sheetsToCreate.length > 0) {
        callbacks.status(`Creating ${sheetsToCreate.length} new sheets...`);
        const createRequests = sheetsToCreate.map(title => ({ addSheet: { properties: { title } } }));
        const createResult = await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests: createRequests } });

        const headerRequests: any[] = [];
        createResult.result.replies?.forEach((reply: any) => {
            const props = reply.addSheet.properties;
            existingSheets.set(props.title, { sheetId: props.sheetId });
            const headers = props.title === LOG_SHEET_NAME ? LOG_HEADERS : MONTHLY_HEADERS;
            headerRequests.push(...generateHeaderRequests(props.sheetId, headers));
        });

        if (headerRequests.length > 0) {
            callbacks.status('Applying headers and styles to new sheets...');
            await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests: headerRequests } });
        }
    }

    const sheetState = new Map<string, { sheetId: number, lastRow: number }>();
    const valueRangesToGet = [...neededSheetNames].filter(name => existingSheets.has(name)).map(name => `${name}!A:A`);
    if (valueRangesToGet.length > 0) {
        const sheetValues = await gapi.client.sheets.spreadsheets.values.batchGet({ spreadsheetId, ranges: valueRangesToGet });
        sheetValues.result.valueRanges?.forEach((rangeResult: any) => {
            const sheetName = rangeResult.range.split('!')[0].replace(/'/g, '');
            const lastRow = rangeResult.values?.length || 0;
            if (existingSheets.has(sheetName)) {
                sheetState.set(sheetName, { sheetId: existingSheets.get(sheetName)!.sheetId, lastRow });
            }
        });
    }

    // --- Phase 4: Build master batch request ---
    const masterRequest: any[] = [];
    for (const [sheetName, rows] of dataBySheet.entries()) {
        const state = sheetState.get(sheetName);
        if (!state) continue;
        masterRequest.push({ appendDimension: { sheetId: state.sheetId, dimension: 'ROWS', length: rows.length } });
        for (let i = 0; i < rows.length; i++) {
            masterRequest.push(...generateRowRequests(state.sheetId, state.lastRow + i, rows[i]));
        }
    }

    // --- Phase 5: Execute ---
    if (masterRequest.length > 0) {
        callbacks.status('Executing a single batch update...');
        callbacks.updateTasksBulk(successfulIds, { status: 'writing' });
        await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests: masterRequest } });
    }

    // --- Phase 6: Log & finalize ---
    const logRows = successfulResults.map(res => [new Date().toISOString(), res.data.messageId, 'Processed', res.data.parsedData.organization, res.data.parsedData.title]);
    if (logRows.length > 0) {
        await gapi.client.sheets.spreadsheets.values.append({ spreadsheetId, range: LOG_SHEET_NAME, valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS', resource: { values: logRows } });
    }
    callbacks.updateTasksBulk(successfulIds, { status: 'done' });
    callbacks.status(`Process complete. ${successfulResults.length} new emails were successfully processed.`);

    return { processed: successfulResults.length };
}

export const SheetsConfig = { LOG_SHEET_NAME, LOG_HEADERS, MONTHLY_HEADERS };
