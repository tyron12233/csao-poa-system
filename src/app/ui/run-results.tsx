import React from 'react';
import type { EmailTask, EmailTaskStatus } from '../../lib/process-emails';
import { Button } from './button';
import { STATUS_LABELS } from './status-pill';
import { TaskItem } from './task-item';

const FINISHED_STATUSES: EmailTaskStatus[] = ['done', 'skipped', 'held', 'error'];
const SUMMARY_STYLES: Partial<Record<EmailTaskStatus, string>> = {
    done: 'text-brand-700 bg-brand-50 ring-brand-200',
    skipped: 'text-gray-700 bg-gray-50 ring-gray-200',
    held: 'text-amber-800 bg-amber-50 ring-amber-200',
    error: 'text-red-700 bg-red-50 ring-red-200',
};

interface RunResultsProps {
    tasks: EmailTask[];
    isProcessing: boolean;
    statusMessage: string;
    statusFilter: EmailTaskStatus | null;
    onFilterChange: (status: EmailTaskStatus | null) => void;
    sheetUrl: string;
    taskAction: (task: EmailTask) => { label: string; onClick: () => void } | undefined;
}

// Progress while a run is going, then a summary with filterable counts and the task list.
export const RunResults: React.FC<RunResultsProps> = ({ tasks, isProcessing, statusMessage, statusFilter, onFilterChange, sheetUrl, taskAction }) => {
    const counts = tasks.reduce<Partial<Record<EmailTaskStatus, number>>>((acc, t) => {
        acc[t.status] = (acc[t.status] || 0) + 1; return acc;
    }, {});
    const finishedCount = tasks.filter((t) => FINISHED_STATUSES.includes(t.status)).length;
    const progressPct = tasks.length ? Math.round((finishedCount / tasks.length) * 100) : 0;
    const visibleTasks = statusFilter ? tasks.filter((t) => t.status === statusFilter) : tasks;

    return (
        <section className="space-y-4" aria-live="polite">
            {isProcessing ? (
                <div className="rounded-xl border border-gray-200 px-4 py-3">
                    <div className="flex justify-between text-xs text-gray-600 mb-2">
                        <span>{tasks.length ? `${finishedCount} of ${tasks.length} emails` : 'Searching Gmail…'}</span>
                        {tasks.length > 0 && <span>{progressPct}%</span>}
                    </div>
                    <div className="h-2 rounded-full bg-gray-100 overflow-hidden">
                        <div className="h-full bg-brand-600 transition-all duration-300" style={{ width: `${progressPct}%` }} />
                    </div>
                    {statusMessage && <p className="mt-2 text-xs text-gray-500 whitespace-pre-line">{statusMessage}</p>}
                </div>
            ) : (
                <div className="rounded-xl border border-gray-200 px-4 py-3 space-y-3">
                    <div className="flex flex-wrap items-center gap-2">
                        <h2 className="text-sm font-semibold text-gray-900 mr-1">Results</h2>
                        {tasks.length === 0 && <span className="text-sm text-gray-500">No new emails found.</span>}
                        {FINISHED_STATUSES.filter((k) => counts[k]).map((k) => {
                            const selected = statusFilter === k;
                            return (
                                <button
                                    key={k}
                                    type="button"
                                    onClick={() => onFilterChange(selected ? null : k)}
                                    aria-pressed={selected}
                                    title={selected ? 'Show all' : `Show only ${STATUS_LABELS[k].toLowerCase()}`}
                                    className={`text-xs font-semibold px-2.5 py-1 rounded-full ring-1 ring-inset transition ${SUMMARY_STYLES[k]} ${selected ? 'ring-2 ring-current' : 'hover:brightness-95'}`}
                                >
                                    {counts[k]} {STATUS_LABELS[k].toLowerCase()}
                                </button>
                            );
                        })}
                        {sheetUrl && counts.done ? (
                            <a href={sheetUrl} target="_blank" rel="noopener noreferrer" className="ml-auto text-sm font-semibold text-brand-700 hover:text-brand-800">
                                Open sheet →
                            </a>
                        ) : null}
                    </div>
                    {statusMessage && (
                        <details className="text-xs text-gray-500">
                            <summary className="cursor-pointer select-none">Run log</summary>
                            <p className="mt-2 whitespace-pre-line">{statusMessage}</p>
                        </details>
                    )}
                </div>
            )}

            {!isProcessing && counts.held ? (
                <div className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">
                    <p className="font-semibold">{counts.held} {counts.held === 1 ? 'email needs' : 'emails need'} dates</p>
                    <p className="text-xs mt-0.5">The activity dates couldn&apos;t be read. Click <span className="font-semibold">Set dates</span> on each one to add it to the sheet.</p>
                </div>
            ) : null}

            {tasks.length > 0 && (
                <div>
                    {statusFilter && (
                        <div className="flex items-center justify-between mb-1.5 text-xs text-gray-500">
                            <span>Showing {STATUS_LABELS[statusFilter].toLowerCase()} only</span>
                            <Button variant="ghost" size="sm" onClick={() => onFilterChange(null)}>Show all</Button>
                        </div>
                    )}
                    <ul className="rounded-xl border border-gray-200 divide-y divide-gray-100 max-h-[28rem] overflow-auto">
                        {visibleTasks.map((task) => (
                            <TaskItem key={task.id} task={task} action={taskAction(task)} />
                        ))}
                    </ul>
                </div>
            )}
        </section>
    );
};
