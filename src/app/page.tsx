"use client";

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { GoogleOAuthProvider, googleLogout, useGoogleLogin } from '@react-oauth/google';
import { processEmails as runProcessEmails } from '../lib/process-emails';
import type { EmailTask } from '../lib/process-emails';

// Import reusable UI components
import { Button } from './ui/button';
import { Input } from './ui/input';
import { Card } from './ui/card';
import { UserProfile } from './ui/user-profile';
import { TaskItem } from './ui/task-item';
import { Modal } from './ui/modal';

let gapi: any = undefined;

interface UserProfile {
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
interface StoredAuth {
  accessToken: string;
  expiresAt: number; // epoch ms
  user: UserProfile;
}

function App() {
  return (
    <GoogleOAuthProvider clientId={GOOGLE_CLIENT_ID}>
      <Home />
    </GoogleOAuthProvider>
  );
}

const Home: React.FC = () => {
  const [user, setUser] = useState<UserProfile | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [tasks, setTasks] = useState<EmailTask[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
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
  const [uploadFolderId, setUploadFolderId] = useState<string>('');
  const [uploadFolderName, setUploadFolderName] = useState<string>('');
  const [isFolderDialogOpen, setIsFolderDialogOpen] = useState(false);
  const [driveFolders, setDriveFolders] = useState<{ id: string; name: string; modifiedTime?: string }[]>([]);
  const [isLoadingFolders, setIsLoadingFolders] = useState(false);
  const [newFolderName, setNewFolderName] = useState('');
  const [folderError, setFolderError] = useState<string | null>(null);
  const [acadStartMonth, setAcadStartMonth] = useState<number>(7); // default July
  const [acadStartYear, setAcadStartYear] = useState<number>(new Date().getFullYear());
  const [manualStartDate, setManualStartDate] = useState<string>('');
  const refreshTimeoutRef = useRef<NodeJS.Timeout | null>(null);

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
        setStatusMessage('Session expiring soon – please click the sign in button to refresh.');
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
      const profile: UserProfile = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${res.access_token}` },
      }).then((r) => r.json());
      setUser(profile);
      // Persist token & expiry
      const expiresInSec = res.expires_in || 3600; // default 1h if not provided
      const stored: StoredAuth = {
        accessToken: res.access_token,
        user: profile,
        expiresAt: Date.now() + expiresInSec * 1000,
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
    setError(null);
    try { localStorage.removeItem(TOKEN_STORAGE_KEY); } catch (_) { }
  };

  const handleProcessEmails = useCallback(async () => {
    if (!spreadsheetId) {
      setError('Please enter a Google Sheet ID.');
      return;
    }
    setError(null);
    setStatusMessage('');
    setIsProcessing(true);
    setTasks([]);
    try {
      await runProcessEmails(gapi, {
        spreadsheetId,
        senderEmail: POA_EMAIL_SENDER,
        useDateFilter,
        uploadFolderId: uploadFolderId || undefined,
        academicYearStartMonth: acadStartMonth,
        academicYearStartYear: acadStartYear,
        manualStartDate: manualStartDate || undefined,
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
  }, [spreadsheetId, useDateFilter]);

  // Simulation handler removed per requirement to remove simulate processing feature.

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
      setStatusMessage('Found sheet – retrieving details...');
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
      setStatusMessage('Ready to process.');
    } catch (e: any) {
      setSheetTitle('');
      setStatusMessage('Could not fetch sheet title – ensure you have access.');
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
    setStatusMessage(`Selected: ${name}`);
  };

  const openFolderDialog = async () => {
    if (!gapi || !accessToken) return;
    setIsFolderDialogOpen(true);
    setFolderError(null);
    setIsLoadingFolders(true);
    try {
      const resp: any = await gapi.client.drive.files.list({
        q: "mimeType='application/vnd.google-apps.folder' and trashed=false",
        orderBy: 'modifiedTime desc',
        pageSize: 100,
        fields: 'files(id,name,modifiedTime)'
      });
      setDriveFolders(resp.result?.files || []);
    } catch (e: any) {
      setFolderError(e.result?.error?.message || 'Failed to load folders.');
    } finally {
      setIsLoadingFolders(false);
    }
  };

  const handleCreateFolder = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newFolderName.trim()) return;
    try {
      setFolderError(null);
      const resp: any = await gapi.client.drive.files.create({
        resource: { name: newFolderName.trim(), mimeType: 'application/vnd.google-apps.folder' },
        fields: 'id,name'
      });
      const folder = resp.result;
      setDriveFolders(prev => [folder, ...prev]);
      setUploadFolderId(folder.id);
      setUploadFolderName(folder.name);
      setNewFolderName('');
    } catch (e: any) {
      setFolderError(e.result?.error?.message || 'Failed to create folder.');
    }
  };

  const selectFolder = (id: string, name: string) => {
    setUploadFolderId(id);
    setUploadFolderName(name);
    setIsFolderDialogOpen(false);
    setStatusMessage(`Selected upload folder: ${name}`);
  };

  return (
    <div className="bg-gray-50 min-h-screen flex flex-col items-center justify-center p-4 font-sans">
      <Card>
        <div className="flex justify-between items-start">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">POA Processor</h1>
            <p className="text-gray-500">to Google Sheets</p>
          </div>
          {user && <UserProfile user={user} onLogout={handleLogout} />}
        </div>

        {!user ? (
          <div className="text-center py-16">
            <h2 className="text-xl font-medium text-gray-800 mb-2">Welcome</h2>
            <p className="text-gray-500 mb-6">Sign in to get started.</p>
            <Button onClick={() => login()}>Sign in with Google</Button>
          </div>
        ) : (
          <div className="mt-8 space-y-6">
            <div>
              <label htmlFor="spreadsheetId" className="block font-semibold text-gray-800 mb-2">
                Target Google Sheet
              </label>
              {editingSheet ? (
                <>
                  <div className="flex gap-2">
                    <Input
                      type="text"
                      id="spreadsheetId"
                      value={spreadsheetInput}
                      onChange={handleSpreadsheetInputChange}
                      placeholder="Paste Google Sheet URL (or use Browse)"
                      aria-describedby="sheet-helper"
                    />
                    <Button type="button" onClick={fetchSheets} disabled={isLoadingSheets}>
                      {isLoadingSheets ? 'Loading...' : 'Browse'}
                    </Button>
                  </div>
                  <div id="sheet-helper" className="mt-1 space-y-1">
                    {spreadsheetInput && !spreadsheetId && (
                      <p className="text-xs text-amber-700 bg-amber-100/70 rounded px-2 py-1">Paste a valid Google Sheet link (https://docs.google.com/spreadsheets/d/...).</p>
                    )}
                    {sheetTitleLoading && (
                      <p className="text-xs text-blue-700 bg-blue-100/70 rounded px-2 py-1">Fetching sheet details...</p>
                    )}
                  </div>
                </>
              ) : (
                <div className="flex items-center justify-between gap-4 bg-white border border-gray-200 rounded-xl px-4 py-3 shadow-sm">
                  <div className="min-w-0">
                    <p className="text-sm text-gray-500 leading-tight">Selected Sheet</p>
                    <p className="font-medium text-gray-900 truncate" title={sheetTitle}>{sheetTitle || 'Untitled Sheet'}</p>
                  </div>
                  <div className="flex gap-2">
                    <Button type="button" onClick={() => { setEditingSheet(true); setShowSheetPicker(false); }}>Change</Button>
                  </div>
                </div>
              )}
            </div>

            {showSheetPicker && (
              <div className="border border-gray-200 rounded-2xl p-3 bg-white shadow-lg max-h-72 overflow-auto space-y-2">
                <div className="flex justify-between items-center mb-1">
                  <h4 className="font-semibold text-sm text-gray-800">Your Recent Spreadsheets</h4>
                  <button
                    className="text-xs text-gray-500 hover:text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-400 rounded"
                    onClick={() => setShowSheetPicker(false)}
                  >Close</button>
                </div>
                {sheets.length === 0 && !isLoadingSheets && (
                  <p className="text-xs text-gray-500">No spreadsheets found.</p>
                )}
                <ul className="space-y-1">
                  {sheets.map(s => (
                    <li key={s.id}>
                      <button
                        type="button"
                        onClick={() => handleSelectSheet(s.id, s.name)}
                        className={`group w-full text-left px-3 py-2 rounded-lg border transition-colors text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 ${spreadsheetId === s.id ? 'bg-blue-50 border-blue-300' : 'bg-gray-50 hover:bg-blue-50 border-gray-200'}`}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="font-medium text-gray-900 truncate" title={s.name}>{s.name}</span>
                          {s.modifiedTime && (
                            <span className="text-[10px] text-gray-500 whitespace-nowrap">{new Date(s.modifiedTime).toLocaleDateString()}</span>
                          )}
                        </div>
                      </button>
                    </li>
                  ))}
                </ul>
              </div>
            )}

            <div className="flex items-center space-x-3 mt-4 p-4 bg-gray-100 rounded-2xl">
              <input
                type="checkbox"
                id="dateFilter"
                className="h-5 w-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                checked={useDateFilter}
                onChange={(e) => setUseDateFilter(e.target.checked)}
              />
              <label htmlFor="dateFilter" className="text-sm font-medium text-gray-700">
                Enable Optimization
                <span className="text-gray-500 font-normal ml-1">
                  (Only fetch newer emails)
                </span>
              </label>
            </div>

            <div className="bg-gray-100 rounded-2xl p-4 space-y-2">
              <div className="flex items-center justify-between gap-3">
                <div className="min-w-0">
                  <p className="text-sm font-medium text-gray-700">Upload PDF Folder (optional)</p>
                  {uploadFolderId ? (
                    <p className="text-xs text-green-700 truncate">{uploadFolderName} (set)</p>
                  ) : (
                    <p className="text-xs text-gray-500">No folder selected. PDFs will be skipped.</p>
                  )}
                </div>
                <Button type="button" onClick={openFolderDialog} disabled={!accessToken}>Select</Button>
              </div>
            </div>

            <div className="grid gap-4 sm:grid-cols-3 bg-gray-100 rounded-2xl p-4">
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">Academic Year Start Month</label>
                <select
                  className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm bg-white focus:outline-none focus:ring-2 focus:ring-blue-500"
                  value={acadStartMonth}
                  onChange={(e) => setAcadStartMonth(Number(e.target.value))}
                >
                  {['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].slice(1).map((m, i) => (
                    <option key={m} value={i + 1}>{m}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">Academic Year Start Year</label>
                <Input type="number" value={acadStartYear} onChange={(e) => setAcadStartYear(Number(e.target.value))} placeholder="2025" />
              </div>
              <div>
                <label className="block text-xs font-semibold text-gray-600 mb-1">Manual Earliest Email Date (override)</label>
                <Input type="date" value={manualStartDate} onChange={(e) => setManualStartDate(e.target.value)} />
              </div>
              <div className="sm:col-span-3 text-[11px] text-gray-500 leading-snug">
                The academic year sets how incomplete or ambiguous year values in emails are inferred. Manual date (if set) restricts Gmail query to messages after that day.
              </div>
            </div>

            <div className="mt-4">
              <Button
                onClick={handleProcessEmails}
                disabled={isProcessing || !spreadsheetId}
                fullWidth
              >
                {isProcessing ? 'Processing...' : 'Process New Emails'}
              </Button>
            </div>

            {error && (
              <p className="text-red-600 text-center bg-red-100 p-3 rounded-2xl mt-4">
                Error: {error}
              </p>
            )}
            {statusMessage && (
              <p className="text-gray-700 text-center bg-blue-50 p-3 rounded-2xl mt-4">
                {statusMessage}
              </p>
            )}

            {tasks.length > 0 && (() => {
              const counts = tasks.reduce<Record<string, number>>((acc, t) => {
                acc[t.status] = (acc[t.status] || 0) + 1; return acc;
              }, {});
              const order = ['queued', 'fetching', 'parsing', 'building_request', 'writing', 'done', 'error'];
              return (
                <div className="mt-8 space-y-4">
                  <div className="flex flex-wrap gap-3 items-center">
                    <h3 className="text-lg font-semibold text-gray-800 mr-2">Tasks</h3>
                    {order.filter(k => counts[k]).map(k => (
                      <span key={k} className="text-[11px] uppercase tracking-wide bg-gray-200 text-gray-700 px-2 py-1 rounded-full font-medium">
                        {k.replace('_', ' ')}: {counts[k]}
                      </span>
                    ))}
                    <span className="text-[11px] text-gray-500 ml-auto">Total: {tasks.length}</span>
                  </div>
                  <ul className="flex flex-wrap gap-2">
                    {tasks.map((task: EmailTask) => (
                      <TaskItem key={task.id} task={task} compact />
                    ))}
                  </ul>
                </div>
              );
            })()}
          </div>
        )}
      </Card>
      <Modal open={isFolderDialogOpen} onClose={() => setIsFolderDialogOpen(false)} title="Select Upload Folder" maxWidthClass="max-w-2xl">
        <div className="space-y-6">
          <form onSubmit={handleCreateFolder} className="flex gap-2 items-end flex-wrap">
            <div className="flex-1 min-w-[220px]">
              <label className="block text-xs font-medium text-gray-600 mb-1">New Folder Name</label>
              <Input value={newFolderName} onChange={e => setNewFolderName(e.target.value)} placeholder="e.g., POA PDFs" />
            </div>
            <Button type="submit" disabled={!newFolderName.trim()}>Create</Button>
          </form>
          {folderError && <p className="text-xs text-red-600 bg-red-50 px-2 py-1 rounded">{folderError}</p>}
          <div>
            <p className="text-xs font-medium text-gray-600 mb-2">Existing Folders</p>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 max-h-64 overflow-auto pr-1">
              {isLoadingFolders && <p className="text-sm text-gray-500">Loading folders...</p>}
              {!isLoadingFolders && driveFolders.length === 0 && (
                <p className="text-sm text-gray-500">No folders found.</p>
              )}
              {driveFolders.map(f => (
                <button
                  key={f.id}
                  type="button"
                  onClick={() => selectFolder(f.id, f.name)}
                  className={`group border rounded-xl px-3 py-2 text-left hover:border-blue-400 hover:bg-blue-50 transition-colors ${uploadFolderId === f.id ? 'border-blue-500 bg-blue-50' : 'border-gray-200 bg-white'}`}
                >
                  <p className="font-medium text-sm text-gray-800 truncate" title={f.name}>{f.name}</p>
                  {f.modifiedTime && <p className="text-[10px] text-gray-500 mt-1">Updated {new Date(f.modifiedTime).toLocaleDateString()}</p>}
                  {uploadFolderId === f.id && <p className="text-[10px] text-blue-600 mt-1 font-semibold">Selected</p>}
                </button>
              ))}
            </div>
          </div>
          <div className="flex justify-end gap-3 pt-2">
            <Button type="button" onClick={() => setIsFolderDialogOpen(false)}>Close</Button>
          </div>
        </div>
      </Modal>
    </div>
  );
};

export default App;