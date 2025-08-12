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
    htmlString: string;
}

// --- Types ---
export interface UserProfile { email: string; name: string; picture: string; }
export type EmailTaskStatus = 'queued' | 'fetching' | 'parsing' | 'uploading_pdf' | 'building_request' | 'writing' | 'done' | 'skipped' | 'held' | 'error';
export interface EmailTask { id: string; subject: string; status: EmailTaskStatus; error?: string; }
export type ParsedSuccessResult = { status: 'success', data: { messageId: string, subject: string, parsedData: any, monthSheetNames: string[], rowData: any[] } };
export type ParsedHeldResult = { status: 'held', messageId: string, subject: string, reason: string, parsedData: any };
export type ParsedResult = ParsedSuccessResult | ParsedHeldResult | { status: 'error', messageId: string, error: string };

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

    details.htmlString = htmlBody;

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

// Date parsing without academic year inference (no remapping of years).
const parseDateNoInference = (dateString: string): Date | null => {
    if (!dateString || dateString === 'Not Found') return null;
    const d = new Date(dateString);
    if (isNaN(d.getTime())) return null;
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
    // Build cell values, supporting formulas and hyperlinks
    const values = rowData.map((val, idx) => {
        if (val && typeof val === 'object') {
            // Support explicit formula injection
            if ('formula' in val && typeof (val as any).formula === 'string') {
                return { userEnteredValue: { formulaValue: (val as any).formula } };
            }
            // Support hyperlink shorthand: { hyperlink: { url, text? } }
            if ('hyperlink' in val && (val as any).hyperlink?.url) {
                const url = String((val as any).hyperlink.url).replace(/"/g, '""');
                const text = String((val as any).hyperlink.text ?? (val as any).hyperlink.url).replace(/"/g, '""');
                const formula = `=HYPERLINK("${url}","${text}")`;
                return { userEnteredValue: { formulaValue: formula } };
            }
        }
        // If the value is a URL string in the "Link Of Approved POA" column (index 9), convert to hyperlink formula
        if (idx === 9 && typeof val === 'string' && /^https?:\/\//i.test(val)) {
            const url = val.replace(/"/g, '""');
            const formula = `=HYPERLINK("${url}","PDF LINK")`;
            return { userEnteredValue: { formulaValue: formula } };
        }
        return { userEnteredValue: { stringValue: String(val ?? '') } };
    });

    const requests: any[] = [
        { updateCells: { start: { sheetId, rowIndex, columnIndex: 0 }, rows: [{ values }], fields: 'userEnteredValue' } },
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
async function fetchAndParseEmail(gapi: any, messageId: string, callbacks: Callbacks, uploadFolderId: string | undefined, acadYearStartYear: number, acadYearStartMonth: number): Promise<ParsedResult> {
    try {
        callbacks.updateTask(messageId, { status: 'fetching' });
        const { result } = await gapi.client.gmail.users.messages.get({ userId: 'me', id: messageId });
        const subject = result.payload.headers.find((h: any) => h.name.toLowerCase() === 'subject')?.value || 'No Subject';
        callbacks.updateTask(messageId, { subject, status: 'parsing' });

        const htmlBody = getEmailBodyHtml(result.payload);
        if (!htmlBody) throw new Error('HTML body not found.');

        const parsedData = parseEmailBody(htmlBody);
        // Disable academic year inference: strictly parse dates; if invalid/missing, hold the task.
        const startDate = parseDateNoInference(parsedData.startDate);
        const endDate = parseDateNoInference(parsedData.endDate);
        if (!startDate || !endDate) {
            const missing = [!startDate ? 'Start Date' : null, !endDate ? 'End Date' : null].filter(Boolean).join(' & ');
            callbacks.updateTask(messageId, { status: 'held', error: `${missing || 'Dates'} need review` });
            return { status: 'held', messageId, subject, reason: `${missing || 'Dates'} missing or invalid`, parsedData };
        }

        const monthSheetNames = getMonthsBetweenDates(startDate, endDate).map(formatMonthYear);


        // add the html body as string (url encoded)
        const baseLink = `https://csao-poa.vercel.app/pdf?html=`
        const pdfLink = baseLink + encodeURIComponent(htmlBody);
        const rowData = [
            'UNSET',
            parsedData.organization,
            parsedData.title,
            parsedData.description,
            parsedData.startDate,
            parsedData.endDate,
            parsedData.time,
            parsedData.venue,
            parsedData.type,
            { hyperlink: { url: pdfLink, text: 'PDF LINK' } },
            ''
        ];

        return { status: 'success', data: { messageId, subject, parsedData, monthSheetNames, rowData } };
    } catch (err: any) {
        const errorMessage = err.result?.error?.message || err.message || 'Unknown parsing error.';
        callbacks.updateTask(messageId, { status: 'error', error: errorMessage });
        return { status: 'error', messageId, error: errorMessage };
    }
}

// PDF generation & Drive upload
async function generateAndUploadPdf(gapi: any, htmlBody: string, subject: string, folderId: string): Promise<string> {
    const { jsPDF } = await import('jspdf');
    await import('html2canvas');
    const doc = new jsPDF({ unit: 'pt', format: 'a4' });
    const container = document.createElement('div');
    container.style.position = 'fixed';
    container.style.left = '-20000px';
    container.style.top = '0';
    container.style.width = '800px';
    container.innerHTML = htmlBody;
    document.body.appendChild(container);
    await new Promise<void>((resolve) => {
        // @ts-ignore html plugin available via import
        doc.html(container, { callback: () => resolve(), autoPaging: 'text', width: 550, windowWidth: 800 });
    });
    document.body.removeChild(container);
    const blob = doc.output('blob') as Blob;
    const fileName = (subject || 'POA Email').substring(0, 80).replace(/[^a-zA-Z0-9 _.-]/g, '_') + '.pdf';

    const metadata = { name: fileName, mimeType: 'application/pdf', parents: [folderId] };
    const boundary = '-------314159265358979323846';
    const delimiter = '\r\n--' + boundary + '\r\n';
    const closeDelim = '\r\n--' + boundary + '--';
    const base64Data = await blobToBase64(blob);
    const multipartRequestBody =
        delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(metadata) +
        delimiter + 'Content-Type: application/pdf\r\nContent-Transfer-Encoding: base64\r\n\r\n' + base64Data + closeDelim;

    const resp: any = await gapi.client.request({
        path: '/upload/drive/v3/files',
        method: 'POST',
        params: { uploadType: 'multipart', fields: 'id,webViewLink' },
        headers: { 'Content-Type': 'multipart/related; boundary=' + boundary },
        body: multipartRequestBody,
    });
    return resp.result.webViewLink || '';
}

function blobToBase64(blob: Blob): Promise<string> {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => {
            const dataUrl = reader.result as string;
            const base64 = dataUrl.split(',')[1];
            resolve(base64);
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });
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
    uploadFolderId?: string; // target Drive folder for PDFs
    academicYearStartMonth: number; // 1-12
    academicYearStartYear: number; // e.g. 2025
    manualStartDate?: string; // ISO date (yyyy-mm-dd) to override earliest email fetch
    callbacks: Callbacks;
};

export async function processEmails(gapi: any, params: ProcessEmailsParams): Promise<{ processed: number }> {
    const { spreadsheetId, senderEmail, useDateFilter, uploadFolderId, academicYearStartMonth, academicYearStartYear, manualStartDate, callbacks } = params;

    // --- Phase 1: Fetch emails ---
    let gmailQuery = `from:${senderEmail} "POA" "The request is now complete."`;
    const processedIds = new Set<string>();

    if (manualStartDate) {
        // Manual start date takes precedence; convert to unix timestamp (start of day)
        const start = new Date(manualStartDate + 'T00:00:00');
        if (!isNaN(start.getTime())) {
            const unix = Math.floor(start.getTime() / 1000);
            gmailQuery += ` after:${unix}`;
            callbacks.status(`Filtering emails after manual start date ${start.toLocaleDateString()}`);
        }
    } else if (useDateFilter) {
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

    // --- Pagination: fetch all messages (up to a max) ---
    const MAX_MESSAGES = 1000;
    const allMessages: any[] = [];
    let pageToken: string | undefined = undefined;
    do {
        const listRes: any = await gapi.client.gmail.users.messages.list({ userId: 'me', q: gmailQuery, maxResults: 500, pageToken });
        const pageMessages = (listRes.result.messages || []).filter((msg: any) => !processedIds.has(msg.id!));
        allMessages.push(...pageMessages);
        callbacks.status(`Fetched ${allMessages.length} candidate emails...`);
        pageToken = listRes.result.nextPageToken;
        if (allMessages.length >= MAX_MESSAGES) {
            callbacks.status(`Reached max limit of ${MAX_MESSAGES} emails; stopping pagination.`);
            break;
        }
    } while (pageToken);

    if (allMessages.length === 0) {
        callbacks.status('No new emails to process.');
        return { processed: 0 };
    }

    // Queue all tasks initially
    callbacks.queue(allMessages.map((msg: any) => ({ id: msg.id!, subject: 'In queue...', status: 'queued' })));

    let totalProcessed = 0;
    const BATCH_SIZE = 10;
    const totalBatches = Math.ceil(allMessages.length / BATCH_SIZE);

    for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
        const start = batchIndex * BATCH_SIZE;
        const batchMessages = allMessages.slice(start, start + BATCH_SIZE);
        callbacks.status(`Processing batch ${batchIndex + 1} of ${totalBatches} (${batchMessages.length} emails)...`);

        // Parse (internal small concurrency: chunks of 5)
        const parsedResults: ParsedResult[] = [];
        for (let i = 0; i < batchMessages.length; i += 5) {
            const chunk = batchMessages.slice(i, i + 5);
            parsedResults.push(...await Promise.all(chunk.map((msg: any) => fetchAndParseEmail(gapi, msg.id!, callbacks, uploadFolderId, academicYearStartYear, academicYearStartMonth))));
        }
        const successfulResults = parsedResults.filter((p): p is ParsedSuccessResult => p.status === 'success');
        const heldResults = parsedResults.filter((p): p is ParsedHeldResult => p.status === 'held');
        const successfulIds = successfulResults.map(p => p.data.messageId);
        const heldIds = heldResults.map(p => p.messageId);
        callbacks.updateTasksBulk(successfulIds, { status: 'building_request' });
        if (heldIds.length) {
            callbacks.updateTasksBulk(heldIds, { status: 'held' });
        }

        // Pre-compute sheet state for this batch
        const dataBySheet = new Map<string, any[][]>();
        const neededSheetNames = new Set<string>([LOG_SHEET_NAME]);
        for (const result of successfulResults) {
            for (const sheetName of result.data.monthSheetNames) {
                if (!dataBySheet.has(sheetName)) dataBySheet.set(sheetName, []);
                dataBySheet.get(sheetName)!.push(result.data.rowData);
                neededSheetNames.add(sheetName);
            }
        }

        // Ensure sheets exist (fetch fresh each batch to include prior creations)
        const initialSheetProps = await gapi.client.sheets.spreadsheets.get({ spreadsheetId });
        const existingSheets = new Map<string, any>(initialSheetProps.result.sheets?.map((s: any) => [s.properties.title, { sheetId: s.properties.sheetId }]));
        const sheetsToCreate = [...neededSheetNames].filter(name => !existingSheets.has(name));

        if (sheetsToCreate.length > 0) {
            callbacks.status(`Creating ${sheetsToCreate.length} new sheets (batch ${batchIndex + 1})...`);
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

        const masterRequest: any[] = [];
        for (const [sheetName, rows] of dataBySheet.entries()) {
            const state = sheetState.get(sheetName);
            if (!state) continue;
            masterRequest.push({ appendDimension: { sheetId: state.sheetId, dimension: 'ROWS', length: rows.length } });
            for (let i = 0; i < rows.length; i++) {
                masterRequest.push(...generateRowRequests(state.sheetId, state.lastRow + i, rows[i]));
            }
        }

        if (masterRequest.length > 0) {
            callbacks.status(`Writing batch ${batchIndex + 1} of ${totalBatches} to sheet...`);
            callbacks.updateTasksBulk(successfulIds, { status: 'writing' });
            await gapi.client.sheets.spreadsheets.batchUpdate({ spreadsheetId, resource: { requests: masterRequest } });
        }

        const logRows = [
            ...successfulResults.map(res => [new Date().toISOString(), res.data.messageId, 'Processed', res.data.parsedData.organization, res.data.parsedData.title]),
            ...heldResults.map(res => [new Date().toISOString(), res.messageId, 'Held (dates needed)', res.parsedData.organization, res.parsedData.title])
        ];
        if (logRows.length > 0) {
            await gapi.client.sheets.spreadsheets.values.append({ spreadsheetId, range: LOG_SHEET_NAME, valueInputOption: 'USER_ENTERED', insertDataOption: 'INSERT_ROWS', resource: { values: logRows } });
        }
        callbacks.updateTasksBulk(successfulIds, { status: 'done' });
        totalProcessed += successfulResults.length;
        const heldMsg = heldResults.length ? `, held ${heldResults.length} for date review` : '';
        callbacks.status(`Batch ${batchIndex + 1} complete. Processed ${successfulResults.length} emails${heldMsg} (Total so far: ${totalProcessed}).`);
    }

    callbacks.status(`Process complete. ${totalProcessed} new emails were successfully processed in ${Math.ceil(allMessages.length / BATCH_SIZE)} batch(es).`);
    // After all batches, summarize any held items and ask user to specify/fix dates.
    try {
        const logData = await gapi.client.sheets.spreadsheets.values.get({ spreadsheetId, range: `${LOG_SHEET_NAME}!A2:E` });
        const rows: any[] = logData.result.values || [];
        const heldNow = rows.filter(r => r[2] && String(r[2]).toLowerCase().startsWith('held'));
        if (heldNow.length > 0) {
            const list = heldNow.map(r => `• ${r[3] || 'Unknown Org'} — ${r[4] || 'Untitled Activity'} (Message ID: ${r[1]})`).join('\n');
            callbacks.status([
                'Action needed: Some items are on hold due to missing/invalid dates.',
                'Please review and specify the correct Start Date and End Date for each held item:',
                list,
                'Next steps:',
                '- Reply with corrected dates in the original email OR',
                `- Manually add/update the dates for the activity in the appropriate monthly sheet (columns "${MONTHLY_HEADERS[4]}" and "${MONTHLY_HEADERS[5]}").`,
                '- Then re-run “Process Emails” to include them.'
            ].join('\n'));
        }
    } catch { /* non-blocking */ }
    return { processed: totalProcessed };
}

// --- Simulation (no Gmail dependency) ---

export const SheetsConfig = { LOG_SHEET_NAME, LOG_HEADERS, MONTHLY_HEADERS };
