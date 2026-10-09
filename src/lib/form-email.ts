// Generic parsing for formapprovals.com "Request #N Complete" emails.
// Every form type (POA, FCA, Pull-out, Sustainability) shares the same email layout:
// a header (request #, date, form title), an approval history, and a table of label/value rows.

export type ApprovalEntry = { action: string; email: string };
export type FieldLink = { label: string; url: string };
export type ResponseField = { label: string; lines: string[]; links: FieldLink[] };

export interface FormEmail {
    requestNumber: string;
    requestUrl: string;
    requestDate: string;
    headerTitle: string;
    statusLabel: string;
    approvalHistory: ApprovalEntry[];
    fields: ResponseField[];
    adminEmail: string;
}

export const clean = (s: string | null | undefined) => (s || "").replace(/\s+/g, " ").trim();

// Normalized form of a label or title used for matching: lowercase, single spaces, no trailing colons.
export const normalizeLabel = (s: string) => clean(s).toLowerCase().replace(/[:\s]+$/, '');

export function parseHeaderMeta(doc: Document): { requestUrl: string; requestNumber: string; requestDate: string; headerTitle: string; statusLabel: string } {
    const contentWrap = doc.querySelector('td.content-wrap') as Element | null;
    let requestUrl = '';
    let requestNumber = '';
    let requestDate = '';
    if (contentWrap) {
        const a = contentWrap.querySelector('a[href]') as Element | null;
        if (a) {
            requestUrl = (a.getAttribute('href') || '');
            requestNumber = clean(a.textContent || '').replace(/^#/, '');
        }
        const topText = clean((contentWrap.textContent || '').split('\n')[0] || '');
        const m = topText.match(/#\d+\s*\|\s*(.+)$/);
        if (m) requestDate = clean(m[1]);
        if (!requestDate) {
            // Header and body may be on one line; grab just the "OCT 06, 2026" part.
            const m2 = clean(contentWrap.textContent).match(/#\d+\s*\|\s*([A-Za-z]{3,9}\.? \d{1,2},? \d{4})/);
            if (m2) requestDate = m2[1];
        }
    }
    const headerTitle = clean((doc.querySelector('#title') as Element | null)?.textContent || '');

    // Prefer the bold word in the sentence "The request is now <strong>Complete</strong>."
    const statusStrong = doc.querySelector('.content-wrap p strong') as Element | null;
    const statusLabel = clean(statusStrong?.textContent || '');

    return { requestUrl, requestNumber, requestDate, headerTitle, statusLabel };
}

export function parseApprovalHistory(doc: Document): ApprovalEntry[] {
    const list: ApprovalEntry[] = [];
    const rows = Array.from(doc.querySelectorAll('.approval-history table tbody tr')) as Element[];
    rows.forEach(row => {
        const span = row.querySelector('span');
        const a = row.querySelector('a.no-link');
        const action = clean(span?.textContent || '');
        const email = clean(a?.textContent || '');
        if (action && email) list.push({ action, email });
    });
    return list;
}

// Split a value cell into display lines, honouring <br> and block elements.
function cellLines(cell: Element): string[] {
    const html = (cell.innerHTML || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/(p|div|li)>/gi, '\n');
    const text = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html').body.textContent || '';
    return text.split('\n').map(clean).filter(Boolean);
}

export function parseFormEmail(htmlBody: string): FormEmail {
    const doc = new DOMParser().parseFromString(htmlBody, 'text/html');
    const header = parseHeaderMeta(doc);

    const fields: ResponseField[] = [];
    const rows = Array.from(doc.querySelectorAll('.response-items tr')) as Element[];
    rows.forEach(row => {
        const tds = Array.from(row.querySelectorAll('td')) as Element[];
        if (tds.length < 2) return;
        const label = clean(tds[0].textContent);
        if (!label) return;
        const links = (Array.from(tds[1].querySelectorAll('a[href]')) as Element[])
            .map(a => ({ label: clean(a.textContent), url: a.getAttribute('href') || '' }))
            .filter(l => l.url);
        fields.push({ label, lines: cellLines(tds[1]), links });
    });

    const adminLink = doc.querySelector('#footer a[href^="mailto:"]') as Element | null;
    const adminEmail = adminLink
        ? clean((adminLink.getAttribute('href') || '').replace(/^mailto:/, '')) || clean(adminLink.textContent)
        : '';

    return {
        ...header,
        statusLabel: header.statusLabel || 'Complete',
        approvalHistory: parseApprovalHistory(doc),
        fields,
        adminEmail,
    };
}

// Reads fields by label and remembers which ones were used, so leftovers can be collected.
export type LabelMatcher = string | RegExp;

export class FieldReader {
    private used = new Set<ResponseField>();
    constructor(private fields: ResponseField[]) { }

    get(...matchers: LabelMatcher[]): ResponseField | undefined {
        for (const m of matchers) {
            const found = this.fields.find(f => {
                const label = normalizeLabel(f.label);
                return typeof m === 'string' ? label === normalizeLabel(m) : m.test(label);
            });
            if (found) { this.used.add(found); return found; }
        }
        return undefined;
    }

    all(matcher: RegExp): ResponseField[] {
        const found = this.fields.filter(f => matcher.test(normalizeLabel(f.label)));
        found.forEach(f => this.used.add(f));
        return found;
    }

    text(...matchers: LabelMatcher[]): string {
        return (this.get(...matchers)?.lines || []).join('\n');
    }

    // Requestor-style fields hold a mailto link; fall back to the visible text.
    email(...matchers: LabelMatcher[]): string {
        const f = this.get(...matchers);
        const mail = f?.links.find(l => l.url.startsWith('mailto:'));
        return mail ? mail.url.replace(/^mailto:/, '') : (f?.lines || []).join(' ');
    }

    links(...matchers: LabelMatcher[]): FieldLink[] {
        return (this.get(...matchers)?.links || []).filter(l => /^https?:\/\//i.test(l.url));
    }

    unused(): ResponseField[] {
        return this.fields.filter(f => !this.used.has(f));
    }
}

// --- Printable view URL (/pdf/form) ---
export function buildFormPdfUrl(baseUrl: string, email: FormEmail): string {
    const url = new URL(baseUrl);
    url.searchParams.set('requestNumber', email.requestNumber);
    url.searchParams.set('requestUrl', email.requestUrl);
    url.searchParams.set('requestDate', email.requestDate);
    url.searchParams.set('headerTitle', email.headerTitle);
    url.searchParams.set('statusLabel', email.statusLabel);
    url.searchParams.set('approvalHistory', JSON.stringify(email.approvalHistory));
    url.searchParams.set('fields', JSON.stringify(email.fields));
    url.searchParams.set('adminEmail', email.adminEmail);
    return url.toString();
}

export function formEmailFromSearchParams(params: URLSearchParams): FormEmail {
    const json = <T,>(key: string, fallback: T): T => {
        try { return JSON.parse(params.get(key) || '') as T; } catch { return fallback; }
    };
    return {
        requestNumber: params.get('requestNumber') || '',
        requestUrl: params.get('requestUrl') || '',
        requestDate: params.get('requestDate') || '',
        headerTitle: params.get('headerTitle') || '',
        statusLabel: params.get('statusLabel') || 'Complete',
        approvalHistory: json<ApprovalEntry[]>('approvalHistory', []),
        fields: json<ResponseField[]>('fields', []),
        adminEmail: params.get('adminEmail') || '',
    };
}
