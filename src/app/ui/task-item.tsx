import React from 'react';
import type { EmailTask } from '../../lib/process-emails';
import { StatusPill } from './status-pill';
import { Button } from './button';

interface TaskItemProps {
    task: EmailTask;
    action?: { label: string; onClick: () => void };
}

export const TaskItem: React.FC<TaskItemProps> = ({ task, action }) => {
    // Every form email has the subject "Request #N Complete", so lead with organization and title.
    const heading = task.organization && task.title
        ? `${task.organization} — ${task.title}`
        : task.title || task.organization || task.subject;
    const meta = [task.requestNumber ? `Request #${task.requestNumber}` : null, task.error].filter(Boolean).join(' · ');

    return (
        <li className="flex items-center gap-3 px-3 py-2.5">
            <div className="min-w-0 flex-1">
                <p className="text-sm font-medium text-gray-800 line-clamp-2 sm:line-clamp-1 break-words" title={heading}>{heading}</p>
                {meta && (
                    <p className={`text-xs truncate ${task.status === 'error' ? 'text-red-600' : 'text-gray-500'}`} title={meta}>{meta}</p>
                )}
            </div>
            <StatusPill status={task.status} />
            {action && (
                <Button variant="secondary" size="sm" onClick={action.onClick}>{action.label}</Button>
            )}
        </li>
    );
};
