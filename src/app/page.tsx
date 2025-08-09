"use client";

import React, { useState, useEffect, useCallback } from 'react';
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
const SCOPES = 'https://www.googleapis.com/auth/gmail.readonly https://www.googleapis.com/auth/spreadsheets';
const POA_EMAIL_SENDER = 'csao.poa@dlsl.edu.ph';

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
  const [statusMessage, setStatusMessage] = useState('');
  const [useDateFilter, setUseDateFilter] = useState(true);

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
      });
    });
    return () => {
      isMounted = false;
    };
  }, [accessToken]);

  const login = useGoogleLogin({
    onSuccess: async (res: any) => {
      setAccessToken(res.access_token);
      const p = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
        headers: { Authorization: `Bearer ${res.access_token}` },
      }).then((r) => r.json());
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
                Target Google Sheet ID:
              </label>
              <Input
                type="text"
                id="spreadsheetId"
                value={spreadsheetId}
                onChange={(e) => setSpreadsheetId(e.target.value)}
                placeholder="Enter Google Sheet ID"
              />
            </div>

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