"use client";

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { GoogleOAuthProvider, googleLogout, useGoogleLogin } from '@react-oauth/google';
import { processEmails as runProcessEmails, reprocessHeldEmailWithDates, getEmailBodyHtml } from '../lib/process-emails';
import type { EmailTask, EmailTaskStatus } from '../lib/process-emails';
import { FORM_TYPE_OPTIONS, formTypeShortLabel } from '../lib/form-types';
import type { FormTypeId } from '../lib/form-types';

// Import reusable UI components
import { Button } from './ui/button';
import { Input } from './ui/input';
import { Card } from './ui/card';
import { UserProfile } from './ui/user-profile';
import { Modal } from './ui/modal';
import { RunResults } from './ui/run-results';

let gapi: any = undefined;

interface GoogleUser {
  email: string;
  name: string;
  picture: string;
}

const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID!;
const API_KEY = process.env.NEXT_PUBLIC_GOOGLE_API_KEY!;
// Added Drive metadata scope so we can list user spreadsheets for a picker
const SCOPES = 'https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.metadata.readonly';
const POA_EMAIL_SENDER = 'csao.poa@dlsl.edu.ph';
const TOKEN_STORAGE_KEY = 'poa_google_auth_v1';
// Each form type writes to its own spreadsheet; remember the last one picked per type.
const SHEET_MEMORY_KEY = 'poa_sheet_by_type_v1';
type SheetMemory = { formType?: FormTypeId; sheets: Partial<Record<FormTypeId, { id: string; name: string }>> };
const readSheetMemory = (): SheetMemory => {
  try {
    const parsed = JSON.parse(localStorage.getItem(SHEET_MEMORY_KEY) || '');
    return { formType: parsed.formType, sheets: parsed.sheets || {} };
  } catch {
    return { sheets: {} };
  }
};
const writeSheetMemory = (memory: SheetMemory) => {
  try { localStorage.setItem(SHEET_MEMORY_KEY, JSON.stringify(memory)); } catch (_) { }
};
interface StoredAuth {
  accessToken: string;
  refreshToken: string | null;
  expiresAt: number; // epoch ms
  user: GoogleUser;
}

const SheetIcon = () => (
  <svg viewBox="0 0 24 24" className="h-9 w-9 shrink-0" aria-hidden="true">
    <rect x="3" y="2" width="18" height="20" rx="2.5" className="fill-brand-600" />
    <path d="M7 8h10M7 12h10M7 16h10M11 8v8" className="stroke-white" strokeWidth="1.5" fill="none" />
  </svg>
);

const Spinner = () => (
  <svg viewBox="0 0 24 24" className="h-4 w-4 animate-spin" aria-hidden="true">
    <circle cx="12" cy="12" r="9" stroke="currentColor" strokeOpacity="0.25" strokeWidth="3" fill="none" />
    <path d="M21 12a9 9 0 0 0-9-9" stroke="currentColor" strokeWidth="3" fill="none" strokeLinecap="round" />
  </svg>
);

function App() {
  return (
    <GoogleOAuthProvider clientId={GOOGLE_CLIENT_ID}>
      <Home />
    </GoogleOAuthProvider>
  );
}

const Home: React.FC = () => {
  const [user, setUser] = useState<GoogleUser | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [tasks, setTasks] = useState<EmailTask[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasRun, setHasRun] = useState(false);
  const [formType, setFormType] = useState<FormTypeId>('poa');
  const [spreadsheetId, setSpreadsheetId] = useState('');
  const [spreadsheetInput, setSpreadsheetInput] = useState('');
  const [statusMessage, setStatusMessage] = useState('');
  const [useDateFilter, setUseDateFilter] = useState(true);
  const [showSheetPicker, setShowSheetPicker] = useState(false);
  const [isLoadingSheets, setIsLoadingSheets] = useState(false);
  const [sheets, setSheets] = useState<{ id: string; name: string; modifiedTime?: string }[]>([]);
  const [sheetTitle, setSheetTitle] = useState<string>('');
  const [sheetTitleLoading, setSheetTitleLoading] = useState(false);
  const [editingSheet, setEditingSheet] = useState(true);
  const [manualStartDate, setManualStartDate] = useState<string>('');
  const refreshTimeoutRef = useRef<NodeJS.Timeout | null>(null);
  // Email details modal state (errors and skipped emails)
  const [isErrorDetailsOpen, setIsErrorDetailsOpen] = useState(false);
  const [isLoadingErrorDetails, setIsLoadingErrorDetails] = useState(false);
  const [detailsTask, setDetailsTask] = useState<EmailTask | null>(null);
  const [errorEmail, setErrorEmail] = useState<{ subject: string; from?: string; date?: string; html?: string; snippet?: string } | null>(null);
  // Held-date modal state
  const [isHeldDialogOpen, setIsHeldDialogOpen] = useState(false);
  const [heldTask, setHeldTask] = useState<EmailTask | null>(null);
  const [heldStartDate, setHeldStartDate] = useState<string>('');
  const [heldEndDate, setHeldEndDate] = useState<string>('');
  const [heldError, setHeldError] = useState<string | null>(null);
  const [heldSubmitting, setHeldSubmitting] = useState(false);
  const [isLoadingHeldPreview, setIsLoadingHeldPreview] = useState(false);
  const [heldPreviewHtml, setHeldPreviewHtml] = useState<string | null>(null);
  const [heldPreviewSnippet, setHeldPreviewSnippet] = useState<string | null>(null);
  // Tasks filter state
  const [statusFilter, setStatusFilter] = useState<EmailTaskStatus | null>(null);

  // Switch form type and restore the spreadsheet last used for it
  const applyFormType = useCallback((type: FormTypeId) => {
    const memory = readSheetMemory();
    writeSheetMemory({ ...memory, formType: type });
    setFormType(type);
    setTasks([]);
    setHasRun(false);
    setStatusFilter(null);
    setStatusMessage('');
    setError(null);
    setShowSheetPicker(false);
    const saved = memory.sheets[type];
    setSpreadsheetId(saved?.id || '');
    setSpreadsheetInput(saved?.name || '');
    setSheetTitle(saved?.name || '');
    setEditingSheet(!saved);
  }, []);

  const rememberSheet = (id: string, name: string) => {
    const memory = readSheetMemory();
    writeSheetMemory({ ...memory, sheets: { ...memory.sheets, [formType]: { id, name } } });
  };

  useEffect(() => {
    applyFormType(readSheetMemory().formType || 'poa');
  }, [applyFormType]);

  // Load cached token/user on first mount
  useEffect(() => {
    try {
      const raw = typeof window !== 'undefined' ? localStorage.getItem(TOKEN_STORAGE_KEY) : null;
      if (!raw) return;
      const stored: StoredAuth = JSON.parse(raw);
      if (stored.expiresAt > Date.now() + 5_000) { // give 5s leeway
        setAccessToken(stored.accessToken);
        setUser(stored.user);
      } else {
        localStorage.removeItem(TOKEN_STORAGE_KEY);
      }
    } catch (_) {
      // ignore corrupt storage
    }
  }, []);

  // Schedule proactive token refresh (best-effort). Without a refresh token we can only prompt user again.
  useEffect(() => {
    if (!accessToken) {
      if (refreshTimeoutRef.current) {
        clearTimeout(refreshTimeoutRef.current);
        refreshTimeoutRef.current = null;
      }
      return;
    }
    try {
      const raw = localStorage.getItem(TOKEN_STORAGE_KEY);
      if (!raw) return;
      const stored: StoredAuth = JSON.parse(raw);
      const msUntilExpiry = stored.expiresAt - Date.now();
      // Attempt a silent-ish refresh 60s before expiry (will still show popup if needed)
      const refreshIn = Math.max(msUntilExpiry - 60_000, 5_000);
      refreshTimeoutRef.current = setTimeout(() => {
        // We cannot truly refresh silently without a refresh token; notify user.
        setStatusMessage('Session expiring soon – please sign out and sign in again to refresh.');
      }, refreshIn);
    } catch (_) {
      // ignore
    }
    return () => {
      if (refreshTimeoutRef.current) clearTimeout(refreshTimeoutRef.current);
    };
  }, [accessToken]);

  useEffect(() => {
    if (!accessToken) return;
    let isMounted = true;

    async function start() {
      gapi = (await import('gapi-script')).gapi;
    }

    start().then(() => {
      if (!isMounted) return;
      gapi.load('client', async () => {
        gapi.client.setApiKey(API_KEY);
        gapi.client.setToken({ access_token: accessToken });
        await gapi.client.load('https://gmail.googleapis.com/$discovery/rest?version=v1');
        await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4');
        try {
          await gapi.client.load('https://www.googleapis.com/discovery/v1/apis/drive/v3/rest');
        } catch (e) {
          // Drive might fail if prior token lacks new scope; user may need to sign in again
        }
      });
    });
    return () => {
      isMounted = false;
    };
  }, [accessToken]);

  const login = useGoogleLogin({
    onSuccess: async (res: any) => {
      setAccessToken(res.access_token);
      const profile: GoogleUser = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${res.access_token}` },
      }).then((r) => r.json());
      setUser(profile);
      // Persist token & expiry
      const expiresInSec = res.expires_in || 3600; // default 1h if not provided
      const stored: StoredAuth = {
        accessToken: res.access_token,
        user: profile,
        expiresAt: Date.now() + expiresInSec * 1000,
        refreshToken: res.refresh_token || null,  // store refresh token if available
      };
      try {
        localStorage.setItem(TOKEN_STORAGE_KEY, JSON.stringify(stored));
      } catch (_) {
        // storage might fail (private mode etc.)
      }
    },
    onError: () => setError('Login Failed'),
    scope: SCOPES,
    // prompt: 'consent', // uncomment if you need to force refresh/consent
  });

  const handleLogout = () => {
    googleLogout();
    setUser(null);
    setAccessToken(null);
    setTasks([]);
    setHasRun(false);
    setError(null);
    try { localStorage.removeItem(TOKEN_STORAGE_KEY); } catch (_) { }
  };

  const handleProcessEmails = useCallback(async () => {
    if (!spreadsheetId) {
      setError('Please select a Google Sheet first.');
      return;
    }
    setError(null);
    setStatusMessage('Searching Gmail…');
    setIsProcessing(true);
    setHasRun(true);
    setStatusFilter(null);
    setTasks([]);
    try {
      await runProcessEmails(gapi, {
        spreadsheetId,
        senderEmail: POA_EMAIL_SENDER,
        useDateFilter,
        manualStartDate: manualStartDate || undefined,
        formType,
        callbacks: {
          status: (msg: string) => setStatusMessage(msg),
          queue: (newTasks: EmailTask[]) => setTasks(newTasks),
          updateTask: (id: string, patch: Partial<EmailTask>) =>
            setTasks((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t))),
          updateTasksBulk: (ids: string[], patch: Partial<EmailTask>) =>
            setTasks((prev) =>
              prev.map((t) => (ids.includes(t.id) ? { ...t, ...patch } : t))
            ),
        },
      });
    } catch (err: any) {
      setError(err.result?.error?.message || err.message || 'An unexpected error occurred.');
    } finally {
      setIsProcessing(false);
    }
  }, [spreadsheetId, useDateFilter, manualStartDate, formType]);

  const openErrorDetails = async (task: EmailTask) => {
    if (!gapi) return;
    setDetailsTask(task);
    setIsErrorDetailsOpen(true);
    setIsLoadingErrorDetails(true);
    setErrorEmail(null);
    try {
      const { result } = await gapi.client.gmail.users.messages.get({ userId: 'me', id: task.id });
      const headers: Array<{ name: string; value: string }> = result.payload.headers || [];
      const subject = headers.find((h) => h.name.toLowerCase() === 'subject')?.value || task.subject || 'No Subject';
      const from = headers.find((h) => h.name.toLowerCase() === 'from')?.value;
      const date = headers.find((h) => h.name.toLowerCase() === 'date')?.value;
      const html = getEmailBodyHtml(result.payload);
      setErrorEmail({ subject, from, date, html, snippet: result.snippet });
    } catch (e) {
      setErrorEmail({ subject: task.subject || 'No Subject', snippet: 'Failed to load email content.' });
    } finally {
      setIsLoadingErrorDetails(false);
    }
  };

  const openHeldDialog = (task: EmailTask) => {
    setHeldTask(task);
    setHeldStartDate('');
    setHeldEndDate('');
    setHeldError(null);
    setHeldPreviewHtml(null);
    setHeldPreviewSnippet(null);
    setIsHeldDialogOpen(true);
    if (!gapi) return;
    setIsLoadingHeldPreview(true);

    gapi.client.gmail.users.messages.get({ userId: 'me', id: task.id }).then(
      ({ result }: any) => {
        try {
          const html = getEmailBodyHtml(result.payload);
          setHeldPreviewHtml(html || null);
          setHeldPreviewSnippet(result.snippet || null);
        } finally {
          setIsLoadingHeldPreview(false);
        }
      },
      () => {
        setHeldPreviewHtml(null);
        setHeldPreviewSnippet('Failed to load email content.');
        setIsLoadingHeldPreview(false);
      }
    );
  };

  const submitHeldDates = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!heldTask) return;
    if (!spreadsheetId) {
      setHeldError('Please select a target Google Sheet first.');
      return;
    }
    if (!heldStartDate || !heldEndDate) {
      setHeldError('Both Start and End dates are required.');
      return;
    }
    const sd = new Date(heldStartDate);
    const ed = new Date(heldEndDate);
    if (isNaN(sd.getTime()) || isNaN(ed.getTime())) {
      setHeldError('Please enter valid dates.');
      return;
    }
    if (sd > ed) {
      setHeldError('Start Date must be on or before End Date.');
      return;
    }
    setHeldError(null);
    setHeldSubmitting(true);
    try {
      await reprocessHeldEmailWithDates(gapi, {
        spreadsheetId,
        messageId: heldTask.id,
        manualStartDate: heldStartDate,
        manualEndDate: heldEndDate,
        callbacks: {
          status: (msg: string) => setStatusMessage(msg),
          queue: (_tasks: EmailTask[]) => { },
          updateTask: (id: string, patch: Partial<EmailTask>) =>
            setTasks((prev) => prev.map((t) => (t.id === id ? { ...t, ...patch } : t))),
          updateTasksBulk: (ids: string[], patch: Partial<EmailTask>) =>
            setTasks((prev) => prev.map((t) => (ids.includes(t.id) ? { ...t, ...patch } : t))),
        },
      });
      setIsHeldDialogOpen(false);
      setHeldTask(null);
    } catch (err: any) {
      setHeldError(err?.message || 'Failed to reprocess the held task.');
    } finally {
      setHeldSubmitting(false);
    }
  };

  // Helper: parse spreadsheet ID from either raw ID or full URL
  const extractSpreadsheetId = (val: string): string | null => {
    if (!val) return null;
    // If it looks like a full URL
    const urlMatch = val.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch) return urlMatch[1];
    // If it's just plausible ID (usually 40+ chars of allowed set)
    if (/^[a-zA-Z0-9-_]{30,}$/.test(val)) return val;
    return null;
  };

  const handleSpreadsheetInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value.trim();
    setSpreadsheetInput(val);
    const id = extractSpreadsheetId(val);
    if (id) {
      setSpreadsheetId(id);
      fetchSheetTitle(id);
    } else {
      setSpreadsheetId('');
      setSheetTitle('');
    }
  };

  const fetchSheetTitle = async (id: string) => {
    if (!gapi || !accessToken || !id) return;
    setSheetTitleLoading(true);
    try {
      const resp: any = await gapi.client.sheets.spreadsheets.get({ spreadsheetId: id, fields: 'properties.title' });
      const title = resp.result?.properties?.title || '';
      setSheetTitle(title);
      setEditingSheet(false);
      rememberSheet(id, title || id);
      setStatusMessage('');
    } catch (e: any) {
      setSheetTitle('');
      setStatusMessage('Could not open that sheet – make sure your account has access to it.');
    } finally {
      setSheetTitleLoading(false);
    }
  };

  const fetchSheets = async () => {
    if (!gapi || !accessToken) return;
    setIsLoadingSheets(true);
    setError(null);
    try {
      const resp: any = await gapi.client.drive.files.list({
        q: "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false",
        orderBy: 'modifiedTime desc',
        pageSize: 50,
        fields: 'files(id,name,modifiedTime)'
      });
      const items = resp.result?.files || [];
      setSheets(items.map((f: any) => ({ id: f.id, name: f.name, modifiedTime: f.modifiedTime })));
      setShowSheetPicker(true);
    } catch (e: any) {
      setError(e.result?.error?.message || 'Failed to list spreadsheets (try signing out/in to grant Drive access).');
    } finally {
      setIsLoadingSheets(false);
    }
  };

  const handleSelectSheet = (id: string, name: string) => {
    setSpreadsheetId(id);
    setSpreadsheetInput(name);
    setShowSheetPicker(false);
    setSheetTitle(name);
    setEditingSheet(false);
    rememberSheet(id, name);
    setStatusMessage('');
  };

  const shortLabel = formTypeShortLabel(formType);
  const sheetUrl = spreadsheetId ? `https://docs.google.com/spreadsheets/d/${spreadsheetId}` : '';

  const taskAction = (task: EmailTask) => {
    if (task.status === 'held') return { label: 'Set dates', onClick: () => openHeldDialog(task) };
    if (task.status === 'error' || task.status === 'skipped') return { label: 'View', onClick: () => openErrorDetails(task) };
    return undefined;
  };

  return (
    <div className="min-h-screen flex flex-col items-center px-4 py-8 sm:py-14">
      <Card>
        <header className="flex justify-between items-center gap-4">
          <div className="flex items-center gap-3 min-w-0">
            <div className="h-10 w-10 shrink-0 rounded-xl bg-brand-600 text-white grid place-items-center font-bold text-[11px] tracking-tight">CSAO</div>
            <div className="min-w-0">
              <h1 className="text-lg sm:text-xl font-bold text-gray-900 leading-tight">CSAO Form Processor</h1>
              <p className="text-sm text-gray-500">Approved form emails → Google Sheets</p>
            </div>
          </div>
          {user && <UserProfile user={user} onLogout={handleLogout} />}
        </header>

        {!user ? (
          <div className="text-center py-14">
            <h2 className="text-xl font-semibold text-gray-900 mb-2">Welcome</h2>
            <p className="text-gray-500 mb-8 max-w-sm mx-auto">
              Sign in with the CSAO Google account that receives the form approval emails.
            </p>
            <Button size="lg" onClick={() => login()}>Sign in with Google</Button>
            {error && <p className="mt-4 text-sm text-red-600">{error}</p>}
          </div>
        ) : (
          <div className="mt-8 space-y-7">
            {/* Form type */}
            <section>
              <h2 className="text-sm font-semibold text-gray-800 mb-2">Form type</h2>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2" role="radiogroup" aria-label="Form type">
                {FORM_TYPE_OPTIONS.map((o) => {
                  const selected = formType === o.id;
                  return (
                    <button
                      key={o.id}
                      type="button"
                      role="radio"
                      aria-checked={selected}
                      title={o.label}
                      disabled={isProcessing}
                      onClick={() => !selected && applyFormType(o.id)}
                      className={`text-left rounded-xl border px-3 py-2.5 transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-brand-500 disabled:cursor-not-allowed ${selected
                        ? 'border-brand-600 bg-brand-50 ring-1 ring-brand-600'
                        : 'border-gray-200 bg-white hover:border-gray-300 hover:bg-gray-50 disabled:opacity-50'}`}
                    >
                      <span className={`block text-sm font-semibold ${selected ? 'text-brand-800' : 'text-gray-900'}`}>{o.shortLabel}</span>
                      <span className="block text-[11px] text-gray-500 truncate">{o.hint}</span>
                    </button>
                  );
                })}
              </div>
              <p className="mt-2 text-xs text-gray-500">
                {formType === 'poa'
                  ? 'POA rows are written to monthly tabs based on the activity dates.'
                  : `${shortLabel} rows are added to one running list in the selected sheet.`}
              </p>
            </section>

            {/* Target sheet */}
            <section>
              <h2 className="text-sm font-semibold text-gray-800 mb-2">
                <label htmlFor="spreadsheetId">Target sheet for {shortLabel}</label>
              </h2>
              {editingSheet ? (
                <>
                  <div className="flex gap-2">
                    <Input
                      type="text"
                      id="spreadsheetId"
                      value={spreadsheetInput}
                      onChange={handleSpreadsheetInputChange}
                      placeholder="Paste a Google Sheet link"
                      aria-describedby="sheet-helper"
                    />
                    <Button variant="secondary" onClick={fetchSheets} disabled={isLoadingSheets}>
                      {isLoadingSheets ? <Spinner /> : 'Browse'}
                    </Button>
                  </div>
                  <div id="sheet-helper" className="mt-1.5 text-xs">
                    {spreadsheetInput && !spreadsheetId && !sheetTitleLoading && (
                      <p className="text-amber-700">Paste a full Google Sheet link (https://docs.google.com/spreadsheets/d/…).</p>
                    )}
                    {sheetTitleLoading && <p className="text-gray-500">Checking sheet…</p>}
                    {!spreadsheetInput && <p className="text-gray-500">Use a separate spreadsheet for each form type.</p>}
                  </div>
                </>
              ) : (
                <div className="flex items-center gap-3 rounded-xl border border-gray-200 bg-white px-3 py-2.5">
                  <SheetIcon />
                  <div className="min-w-0 flex-1">
                    <p className="font-medium text-gray-900 truncate" title={sheetTitle}>{sheetTitle || 'Untitled sheet'}</p>
                    <p className="text-xs text-gray-500">Google Sheet</p>
                  </div>
                  <a
                    href={sheetUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="text-sm font-semibold text-brand-700 hover:text-brand-800 px-2 py-1.5 rounded-lg hover:bg-brand-50"
                  >
                    Open
                  </a>
                  <Button variant="ghost" size="sm" onClick={() => { setEditingSheet(true); setShowSheetPicker(false); }} disabled={isProcessing}>
                    Change
                  </Button>
                </div>
              )}

              {showSheetPicker && (
                <div className="mt-2 border border-gray-200 rounded-xl bg-white shadow-lg overflow-hidden">
                  <div className="flex justify-between items-center px-3 py-2 border-b border-gray-100 bg-gray-50">
                    <p className="text-xs font-semibold text-gray-700">Recent spreadsheets</p>
                    <Button variant="ghost" size="sm" onClick={() => setShowSheetPicker(false)}>Close</Button>
                  </div>
                  {sheets.length === 0 && !isLoadingSheets && (
                    <p className="text-sm text-gray-500 px-3 py-4">No spreadsheets found.</p>
                  )}
                  <ul className="max-h-64 overflow-auto divide-y divide-gray-100">
                    {sheets.map((s) => (
                      <li key={s.id}>
                        <button
                          type="button"
                          onClick={() => handleSelectSheet(s.id, s.name)}
                          className={`w-full text-left px-3 py-2 text-sm flex items-center justify-between gap-2 focus:outline-none focus-visible:bg-brand-50 ${spreadsheetId === s.id ? 'bg-brand-50' : 'hover:bg-gray-50'}`}
                        >
                          <span className="font-medium text-gray-900 truncate" title={s.name}>{s.name}</span>
                          {s.modifiedTime && (
                            <span className="text-[11px] text-gray-500 whitespace-nowrap">{new Date(s.modifiedTime).toLocaleDateString()}</span>
                          )}
                        </button>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </section>

            {/* Options */}
            <details className="group rounded-xl border border-gray-200 bg-gray-50/60">
              <summary className="cursor-pointer select-none list-none px-4 py-3 flex items-center justify-between gap-3 text-sm">
                <span className="font-semibold text-gray-800">Options</span>
                <span className="text-xs text-gray-500 truncate">
                  {manualStartDate
                    ? `Emails since ${new Date(manualStartDate + 'T00:00:00').toLocaleDateString()}`
                    : useDateFilter ? 'Only new emails' : 'All emails'}
                  <span className="ml-2 inline-block transition-transform group-open:rotate-90">›</span>
                </span>
              </summary>
              <div className="px-4 pb-4 space-y-4">
                <label className="flex items-start gap-3 text-sm text-gray-700">
                  <input
                    type="checkbox"
                    className="mt-0.5 h-4 w-4 rounded border-gray-300 accent-brand-600"
                    checked={useDateFilter}
                    onChange={(e) => setUseDateFilter(e.target.checked)}
                  />
                  <span>
                    <span className="font-medium">Only fetch new emails</span>
                    <span className="block text-xs text-gray-500">Skips emails already recorded in the {shortLabel} log.</span>
                  </span>
                </label>
                <div>
                  <label htmlFor="manualStartDate" className="block text-sm font-medium text-gray-700 mb-1">Earliest email date</label>
                  <div className="flex gap-2 items-center max-w-xs">
                    <Input id="manualStartDate" type="date" value={manualStartDate} onChange={(e) => setManualStartDate(e.target.value)} />
                    {manualStartDate && <Button variant="ghost" size="sm" onClick={() => setManualStartDate('')}>Clear</Button>}
                  </div>
                  <p className="mt-1 text-xs text-gray-500">Optional. Only emails received on or after this day are processed (overrides the option above).</p>
                </div>
              </div>
            </details>

            <Button size="lg" fullWidth onClick={handleProcessEmails} disabled={isProcessing || !spreadsheetId}>
              {isProcessing ? (<><Spinner /> Processing…</>) : `Process new ${shortLabel} emails`}
            </Button>

            {error && (
              <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-sm text-red-700" role="alert">
                {error}
              </div>
            )}

            {!hasRun && statusMessage && (
              <p className="text-sm text-gray-600 text-center">{statusMessage}</p>
            )}

            {/* Run progress & results */}
            {hasRun && (
              <RunResults
                tasks={tasks}
                isProcessing={isProcessing}
                statusMessage={statusMessage}
                statusFilter={statusFilter}
                onFilterChange={setStatusFilter}
                sheetUrl={sheetUrl}
                taskAction={taskAction}
              />
            )}
          </div>
        )}
      </Card>
      <Modal open={isErrorDetailsOpen} onClose={() => setIsErrorDetailsOpen(false)} title="Email preview" maxWidthClass="max-w-4xl">
        {detailsTask?.error && (
          <p className={`mb-3 rounded-lg px-3 py-2 text-sm ${detailsTask.status === 'error' ? 'bg-red-50 text-red-700' : 'bg-gray-100 text-gray-700'}`}>
            {detailsTask.status === 'error' ? 'Error: ' : 'Skipped: '}{detailsTask.error}
          </p>
        )}
        {isLoadingErrorDetails && (
          <p className="text-sm text-gray-600">Loading message…</p>
        )}
        {!isLoadingErrorDetails && errorEmail && (
          <div className="space-y-3">
            <div>
              <p className="text-base font-medium text-gray-900">{errorEmail.subject}</p>
              {(errorEmail.from || errorEmail.date) && (
                <p className="text-xs text-gray-500">
                  {errorEmail.from}{errorEmail.from && errorEmail.date ? ' · ' : ''}{errorEmail.date && new Date(errorEmail.date).toLocaleString()}
                </p>
              )}
            </div>
            {errorEmail.html ? (
              <div className="border rounded-lg overflow-hidden">
                <iframe
                  title="Email preview"
                  sandbox=""
                  className="w-full h-[60vh] bg-white"
                  srcDoc={errorEmail.html}
                />
              </div>
            ) : (
              <pre className="text-xs text-gray-700 bg-gray-50 p-3 rounded border overflow-auto max-h-[60vh] whitespace-pre-wrap">
                {errorEmail.snippet || 'No content available.'}
              </pre>
            )}
          </div>
        )}
      </Modal>
      <Modal open={isHeldDialogOpen} onClose={() => setIsHeldDialogOpen(false)} title="Set activity dates" maxWidthClass="max-w-4xl">
        <form onSubmit={submitHeldDates} className="space-y-4">
          <div className="text-sm text-gray-700">
            <p className="font-medium text-gray-900">
              {heldTask?.organization && heldTask?.title ? `${heldTask.organization} — ${heldTask.title}` : heldTask?.title || heldTask?.subject || 'Held email'}
            </p>
            <p className="text-xs text-gray-500">Enter the Start and End dates from the email below; the activity will be written to the matching monthly tab(s).</p>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label htmlFor="heldStartDate" className="block text-xs font-semibold text-gray-600 mb-1">Start date</label>
              <Input id="heldStartDate" type="date" value={heldStartDate} onChange={(e) => setHeldStartDate(e.target.value)} required />
            </div>
            <div>
              <label htmlFor="heldEndDate" className="block text-xs font-semibold text-gray-600 mb-1">End date</label>
              <Input id="heldEndDate" type="date" value={heldEndDate} onChange={(e) => setHeldEndDate(e.target.value)} required />
            </div>
          </div>
          {heldError && <p className="text-xs text-red-600 bg-red-50 px-2 py-1 rounded">{heldError}</p>}
          <div className="space-y-2">
            <p className="text-xs font-medium text-gray-600">Email</p>
            {isLoadingHeldPreview && (
              <p className="text-sm text-gray-600">Loading message…</p>
            )}
            {!isLoadingHeldPreview && (heldPreviewHtml || heldPreviewSnippet) && (
              heldPreviewHtml ? (
                <div className="border rounded-lg overflow-hidden bg-white">
                  <iframe
                    title="Held email preview"
                    sandbox=""
                    className="w-full h-[45vh]"
                    srcDoc={heldPreviewHtml}
                  />
                </div>
              ) : (
                <pre className="text-xs text-gray-700 bg-gray-50 p-3 rounded border overflow-auto max-h-[45vh] whitespace-pre-wrap">
                  {heldPreviewSnippet}
                </pre>
              )
            )}
          </div>
          <div className="flex justify-end gap-2 pt-2">
            <Button variant="secondary" onClick={() => setIsHeldDialogOpen(false)} disabled={heldSubmitting}>Cancel</Button>
            <Button type="submit" disabled={heldSubmitting}>{heldSubmitting ? (<><Spinner /> Saving…</>) : 'Add to sheet'}</Button>
          </div>
        </form>
      </Modal>
    </div>
  );
};

export default App;
