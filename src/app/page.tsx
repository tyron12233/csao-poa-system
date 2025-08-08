"use client";

import React, { useState, useEffect, useCallback } from 'react';
import { GoogleOAuthProvider, googleLogout, useGoogleLogin } from '@react-oauth/google';
import { processEmails as runProcessEmails } from '../lib/process-emails';
import type { EmailTask } from '../lib/process-emails';
import dynamic from 'next/dynamic';


let gapi: any = undefined;


// --- UI Types ---
interface UserProfile { email: string; name: string; picture: string; }

// --- Configuration ---
const GOOGLE_CLIENT_ID = process.env.NEXT_PUBLIC_GOOGLE_CLIENT_ID!;
const API_KEY = process.env.NEXT_PUBLIC_GOOGLE_API_KEY!;
const SCOPES = 'https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/spreadsheets';
const POA_EMAIL_SENDER = 'csao.poa@dlsl.edu.ph';

// --- Main Application Component ---
function App() { return (<GoogleOAuthProvider clientId={GOOGLE_CLIENT_ID}><Home /></GoogleOAuthProvider>); }

const Home: React.FC = () => {
  const [user, setUser] = useState<UserProfile | null>(null);
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [tasks, setTasks] = useState<EmailTask[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [spreadsheetId, setSpreadsheetId] = useState('');
  const [statusMessage, setStatusMessage] = useState('');
  const [useDateFilter, setUseDateFilter] = useState(true);



  useEffect(() => {
    if (!accessToken) return;
    let isMounted = true;

    async function start() {
      gapi = (await import('gapi-script')).default;
    }

    start().then(() => {
      if (!isMounted) return;
      gapi.load('client', async () => {
        gapi.client.setApiKey(API_KEY);
        gapi.client.setToken({ access_token: accessToken });
        await gapi.client.load('https://gmail.googleapis.com/$discovery/rest?version=v1');
        await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4');
      });
    });
    return () => { isMounted = false; };
  }, [accessToken]);

  const login = useGoogleLogin({
    onSuccess: async (res: any) => {
      setAccessToken(res.access_token);
      const p = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: `Bearer ${res.access_token}` } }).then(r => r.json());
      setUser(p);
    },
    onError: () => setError('Login Failed'),
    scope: SCOPES,
  });

  const handleLogout = () => {
    googleLogout();
    setUser(null);
    setAccessToken(null);
    setTasks([]);
    setError(null);
  };

  // --- Business action moved to lib ---
  const handleProcessEmails = useCallback(async () => {
    if (!spreadsheetId) { setError('Please enter a Google Sheet ID.'); return; }
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
            setTasks(prev => prev.map(t => t.id === id ? { ...t, ...patch } : t)),
          updateTasksBulk: (ids: string[], patch: Partial<EmailTask>) =>
            setTasks(prev => prev.map(t => ids.includes(t.id) ? { ...t, ...patch } : t)),
        }
      });
    } catch (err: any) {
      setError(err.result?.error?.message || err.message || 'An unexpected error occurred.');
    } finally {
      setIsProcessing(false);
    }
  }, [spreadsheetId, useDateFilter]);

  return (
    <div className="bg-slate-100 min-h-screen flex flex-col items-center p-4 font-sans">
      <div className="w-full max-w-4xl bg-white rounded-xl shadow-lg p-6 space-y-6">
        <div className="flex justify-between items-center border-b pb-4">
          <h1 className="text-2xl font-bold text-gray-800">POA to Google Sheets Processor</h1>
          {user && (
            <button onClick={handleLogout} className="text-sm bg-red-500 hover:bg-red-600 text-white font-semibold px-3 py-1 rounded-md">
              Logout
            </button>
          )}
        </div>

        {!user ? (
          <div className="text-center py-10">
            <button onClick={() => login()} className="bg-blue-600 hover:bg-blue-700 text-white font-bold px-6 py-3 rounded-lg shadow-md">
              Sign in with Google
            </button>
          </div>
        ) : (
          <div>
            <div>
              <label htmlFor="spreadsheetId" className="block font-semibold text-gray-700">Target Google Sheet ID:</label>
              <input
                type="text"
                id="spreadsheetId"
                value={spreadsheetId}
                onChange={(e) => setSpreadsheetId(e.target.value)}
                placeholder="Enter Google Sheet ID"
                className="w-full p-2 border border-gray-300 rounded-md"
              />
            </div>

            <div className="flex items-center space-x-3 mt-4 p-3 bg-gray-50 rounded-md">
              <input
                type="checkbox"
                id="dateFilter"
                className="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                checked={useDateFilter}
                onChange={(e) => setUseDateFilter(e.target.checked)}
              />
              <label htmlFor="dateFilter" className="text-sm font-medium text-gray-700">
                Enable Optimization <span className="text-gray-500 font-normal">(Only fetch emails newer than the last processed one)</span>
              </label>
            </div>

            <button
              onClick={handleProcessEmails}
              disabled={isProcessing || !spreadsheetId}
              className="w-full mt-4 bg-green-600 text-white px-4 py-3 rounded-lg font-bold text-lg disabled:bg-gray-400 disabled:cursor-not-allowed hover:bg-green-700"
            >
              {isProcessing ? 'Processing...' : 'Process New Emails'}
            </button>

            {error && <p className="text-red-600 text-center bg-red-100 p-3 rounded-md mt-4">Error: {error}</p>}
            {statusMessage && <p className="text-gray-700 text-center bg-blue-50 p-3 rounded-md mt-4">{statusMessage}</p>}

            {tasks.length > 0 && (
              <ul className="mt-6 space-y-3">
                {tasks.map((task: EmailTask) => (
                  <li key={task.id} className="flex justify-between items-center bg-gray-50 p-3 rounded-lg border">
                    <div className="flex-1 min-w-0">
                      <p className="font-medium text-gray-800 truncate">{task.subject}</p>
                      {task.error && <p className="text-xs text-red-500 truncate">Error: {task.error}</p>}
                    </div>
                    <span
                      className={`ml-4 px-3 py-1 rounded-full text-xs font-semibold capitalize ${({ done: 'bg-green-100 text-green-800', error: 'bg-red-100 text-red-800', fetching: 'bg-blue-100 text-blue-800 animate-pulse', parsing: 'bg-yellow-100 text-yellow-800 animate-pulse', building_request: 'bg-indigo-100 text-indigo-800', writing: 'bg-purple-100 text-purple-800 animate-pulse', queued: 'bg-gray-200 text-gray-800', skipped: 'bg-gray-300 text-gray-600' } as any)[task.status] || 'bg-gray-200'
                        }`}
                    >
                      {task.status.replace('_', ' ')}
                    </span>
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default App;