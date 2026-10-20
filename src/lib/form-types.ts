import { FieldReader, FieldLink, FormEmail, normalizeLabel } from './form-email';

export type FormTypeId = 'poa' | 'fca' | 'pullout' | 'sustainability';

// A sheet cell: plain text, or a clickable link rendered via =HYPERLINK().
export type CellValue = string | { hyperlink: { url: string; text?: string } };

export interface ListColumn { header: string; width: number; wrap?: boolean }

export interface ListFormConfig {
    id: Exclude<FormTypeId, 'poa'>;
    label: string;
    shortLabel: string;
    // Matched against the normalized email title (the big heading in the email body).
    titleMatch: RegExp;
    // Phrase added to the Gmail search to narrow results; the title check is what decides.
    gmailPhrase: string;
    sheetName: string;
    logSheetName: string;
    columns: ListColumn[];
    buildRow: (email: FormEmail, pdfUrl: string) => { row: CellValue[]; organization: string; title: string };
}

export const FORM_TYPE_OPTIONS: { id: FormTypeId; label: string; shortLabel: string; hint: string }[] = [
    { id: 'poa', label: 'Program of Activities (POA)', shortLabel: 'POA', hint: 'Program of Activities' },
    { id: 'fca', label: 'Funds Collection Approval (FCA)', shortLabel: 'FCA', hint: 'Funds Collection' },
    { id: 'pullout', label: 'Pull-out Form (Class/Students Participation)', shortLabel: 'Pull-out', hint: 'Class participation' },
    { id: 'sustainability', label: 'Sustainability Checklist and Approval Form', shortLabel: 'Sustainability', hint: 'Checklist & approval' },
];

export const formTypeShortLabel = (id: FormTypeId) => FORM_TYPE_OPTIONS.find(o => o.id === id)!.shortLabel;

const isBlank = (s: string) => /^(n\/?a|none|-+)?$/i.test(s.trim());

const linkCell = (links: FieldLink[]): CellValue => {
    if (links.length === 0) return '';
    if (links.length === 1) return { hyperlink: { url: links[0].url, text: links[0].label || 'File Upload 1' } };
    // HYPERLINK() holds a single URL; list them all as text (Sheets makes them clickable).
    return links.map(l => l.url).join('\n');
};

const pdfCell = (pdfUrl: string): CellValue => ({ hyperlink: { url: pdfUrl, text: 'PDF LINK' } });

// Any rows the column mapping didn't consume (e.g. form sections added later), as "Label: value".
const otherDetails = (reader: FieldReader): string =>
    reader.unused()
        .map(f => `${f.label.replace(/:+\s*$/, '')}: ${f.lines.join(' / ')}`)
        .join('\n');

const OTHER_DETAILS: ListColumn = { header: 'Other Details', width: 300, wrap: true };
const PDF_COLUMN: ListColumn = { header: 'PDF Link', width: 120 };

const FCA: ListFormConfig = {
    id: 'fca',
    label: 'Funds Collection Approval (FCA)',
    shortLabel: 'FCA',
    titleMatch: /funds collection approval/,
    gmailPhrase: 'Funds Collection Approval',
    sheetName: 'FCA',
    logSheetName: 'FCA Log',
    columns: [
        { header: 'Request #', width: 90 },
        { header: 'Date Requested', width: 130 },
        { header: 'Organization', width: 250 },
        { header: 'Name of Activity', width: 280 },
        { header: 'Start of Collection', width: 150 },
        { header: 'End of Collection', width: 150 },
        { header: 'Collection Items', width: 380, wrap: true },
        { header: 'Total Collection', width: 140 },
        { header: 'Approved POA', width: 130 },
        { header: 'Prepared By / Position', width: 260, wrap: true },
        { header: 'Requestor', width: 200 },
        OTHER_DETAILS,
        PDF_COLUMN,
    ],
    buildRow(email, pdfUrl) {
        const r = new FieldReader(email.fields);
        const organization = r.text('Organization', 'Name of Organization');
        const title = r.text('Name of Activity', 'Title of Activity');
        const requestor = r.email('Requestor');
        const startDate = r.text('Start Date of Collection');
        const endDate = r.text('End Date of Collection');

        // Item slots are labelled "A) ITEM/PURPOSE", "A) NUMBER OF PARTICIPANTS", ... for A, B, C, ...
        const letters = r.all(/^[a-z]\)\s*item/).map(f => normalizeLabel(f.label)[0]);
        const items = letters.flatMap(l => {
            const item = r.text(new RegExp(`^${l}\\)\\s*item`));
            const count = r.text(new RegExp(`^${l}\\)\\s*number of participant`));
            const cost = r.text(new RegExp(`^${l}\\)\\s*cost per participant`));
            const total = r.text(new RegExp(`^${l}\\)\\s*total`));
            if (isBlank(item)) return [];
            let line = item;
            if (!isBlank(count) && !isBlank(cost)) line += ` — ${count} × ${cost}`;
            if (!isBlank(total)) line += `${line === item ? ' —' : ' ='} ${total}`;
            return [line];
        });

        const row: CellValue[] = [
            email.requestNumber,
            email.requestDate,
            organization,
            title,
            startDate,
            endDate,
            items.join('\n'),
            r.text('Total Collection'),
            linkCell(r.links(/^upload (the|your) approved poa/)),
            r.text(/^prepared by/),
            requestor,
            otherDetails(r),
            pdfCell(pdfUrl),
        ];
        return { row, organization, title };
    },
};

const PULLOUT: ListFormConfig = {
    id: 'pullout',
    label: 'Pull-out Form (Class/Students Participation)',
    shortLabel: 'Pull-out',
    titleMatch: /participation in school activities/,
    gmailPhrase: 'Participation in School Activities',
    sheetName: 'Pull-out Requests',
    logSheetName: 'Pull-out Log',
    columns: [
        { header: 'Request #', width: 90 },
        { header: 'Date Requested', width: 130 },
        { header: 'Organization', width: 220 },
        { header: 'Purpose', width: 320, wrap: true },
        { header: 'Date', width: 140 },
        { header: 'Time', width: 160 },
        { header: 'Venue', width: 200 },
        { header: 'No. of Students', width: 120 },
        { header: 'Classes / Students Requested', width: 380, wrap: true },
        { header: 'Requested By / Position', width: 260, wrap: true },
        { header: 'Approved POA / Invitation', width: 160 },
        { header: 'Requestor', width: 200 },
        OTHER_DETAILS,
        PDF_COLUMN,
    ],
    buildRow(email, pdfUrl) {
        const r = new FieldReader(email.fields);
        const organization = r.text('Name of Organization', 'Organization');
        const title = r.text('Purpose');
        const requestor = r.email('Requestor');

        // Program headings look like "(Bachelor of Elementary Education)"; everything else is a student.
        const studentLines = r.get(/^classes\/students/)?.lines || [];
        const isGroupHeading = (s: string) => /^\(.*\)$/.test(s);
        const studentCount = studentLines.filter(s => !isGroupHeading(s)).length;
        const students = studentLines
            .map((s, i) => (isGroupHeading(s) && i > 0 ? `\n${s}` : s))
            .join('\n');

        const row: CellValue[] = [
            email.requestNumber,
            email.requestDate,
            organization,
            title,
            r.text('Date'),
            r.text('Time'),
            r.text('Venue'),
            studentCount ? String(studentCount) : '',
            students,
            r.text(/^requested by/),
            linkCell(r.links(/^upload your approved poa/)),
            requestor,
            otherDetails(r),
            pdfCell(pdfUrl),
        ];
        return { row, organization, title };
    },
};

const SUSTAINABILITY: ListFormConfig = {
    id: 'sustainability',
    label: 'Sustainability Checklist and Approval Form',
    shortLabel: 'Sustainability',
    titleMatch: /sustainability checklist/,
    gmailPhrase: 'Sustainability Checklist',
    sheetName: 'Sustainability Checklist',
    logSheetName: 'Sustainability Log',
    columns: [
        { header: 'Request #', width: 90 },
        { header: 'Date Requested', width: 130 },
        { header: 'Event/Project Title', width: 280, wrap: true },
        { header: 'Organization/Office', width: 220 },
        { header: 'Proponent', width: 180 },
        { header: 'Date of Event', width: 130 },
        { header: 'Time', width: 160 },
        { header: 'Venue', width: 180 },
        { header: 'Participants', width: 140 },
        { header: 'Food & Beverage Packaging', width: 200, wrap: true },
        { header: 'Decorations', width: 200, wrap: true },
        { header: 'Marketing Materials', width: 200, wrap: true },
        { header: 'Others', width: 200, wrap: true },
        { header: 'Waste Management Plan', width: 220, wrap: true },
        { header: 'Compliance Commitment', width: 160 },
        { header: 'Date of Commitment', width: 140 },
        { header: 'Material Pictures', width: 140 },
        { header: 'Approved POA / Concept Paper', width: 170 },
        { header: 'Requestor', width: 200 },
        OTHER_DETAILS,
        PDF_COLUMN,
    ],
    buildRow(email, pdfUrl) {
        const r = new FieldReader(email.fields);
        const organization = r.text('Organization/Office', 'Organization');
        const title = r.text('Event/Project Title');
        const requestor = r.email('Requestor');

        const row: CellValue[] = [
            email.requestNumber,
            email.requestDate,
            title,
            organization,
            r.text('Proponent Name'),
            r.text('Date of Event'),
            r.text('Time of Event'),
            r.text('Venue'),
            r.text('Participants'),
            r.text(/food & beverage/),
            r.text(/decorations/),
            r.text(/marketing materials/),
            r.text(/materials usage declaration.*others/),
            r.text(/waste management plan/),
            r.text(/compliance commitment/),
            r.text('Date of Commitment'),
            linkCell(r.links(/^attached pictures/)),
            linkCell(r.links(/^upload your approved poa/)),
            requestor,
            otherDetails(r),
            pdfCell(pdfUrl),
        ];
        return { row, organization, title };
    },
};

export const LIST_FORM_TYPES: Record<ListFormConfig['id'], ListFormConfig> = {
    fca: FCA,
    pullout: PULLOUT,
    sustainability: SUSTAINABILITY,
};

// Every form comes from the same sender with the same subject ("Request #N Complete"),
// so the email's title heading is the only reliable way to tell them apart.
// Anything that isn't one of the known list forms is treated as a POA (the original behaviour).
export function detectFormType(headerTitle: string): FormTypeId {
    const title = normalizeLabel(headerTitle);
    const match = Object.values(LIST_FORM_TYPES).find(c => c.titleMatch.test(title));
    return match ? match.id : 'poa';
}
