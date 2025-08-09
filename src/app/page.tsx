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
      setStatusMessage('Parsed Sheet ID successfully.');
    } else {
      setSpreadsheetId('');
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
    setSpreadsheetInput(id);
    setShowSheetPicker(false);
    setStatusMessage(`Selected sheet: ${name}`);
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
              <label htmlFor="spreadsheetId" className="block font-semibold text-gray-700 mb-2">
                Target Google Sheet:
              </label>
              <div className="flex gap-2">
                <Input
                  type="text"
                  id="spreadsheetId"
                  value={spreadsheetInput}
                  onChange={handleSpreadsheetInputChange}
                  placeholder="Paste Sheet URL or ID"
                />
                <Button type="button" onClick={fetchSheets} disabled={isLoadingSheets}>
                  {isLoadingSheets ? 'Loading...' : 'Browse'}
                </Button>
              </div>
              {spreadsheetInput && !spreadsheetId && (
                <p className="text-xs text-red-600 mt-1">Could not parse a valid Sheet ID.</p>
              )}
              {spreadsheetId && (
                <p className="text-xs text-green-600 mt-1">Using ID: {spreadsheetId}</p>
              )}
            </div>

            {showSheetPicker && (
              <div className="border rounded-2xl p-3 bg-white shadow-sm max-h-64 overflow-auto space-y-2">
                <div className="flex justify-between items-center mb-1">
                  <h4 className="font-semibold text-sm">Your Recent Spreadsheets</h4>
                  <button
                    className="text-xs text-gray-500 hover:text-gray-700"
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
                        className={`w-full text-left px-2 py-1 rounded hover:bg-blue-50 text-sm ${spreadsheetId === s.id ? 'bg-blue-100' : ''}`}
                      >
                        <span className="font-medium">{s.name}</span>
                        {s.modifiedTime && (
                          <span className="ml-2 text-[10px] text-gray-500">{new Date(s.modifiedTime).toLocaleDateString()}</span>
                        )}
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

            <Button
              onClick={handleProcessEmails}
              disabled={isProcessing || !spreadsheetId}
              fullWidth
            >
              {isProcessing ? 'Processing...' : 'Process New Emails'}
            </Button>

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

            {tasks.length > 0 && (
              <div className="mt-8">
                <h3 className="text-lg font-semibold text-gray-800 mb-4">Tasks</h3>
                <ul className="space-y-3">
                  {tasks.map((task: EmailTask) => (
                    <TaskItem key={task.id} task={task} />
                  ))}
                </ul>
              </div>
            )}
          </div>
        )}
      </Card>
    </div>
  );
};

export default App;