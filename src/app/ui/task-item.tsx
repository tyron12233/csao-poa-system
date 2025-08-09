
import React from 'react';
import type { EmailTask } from '../../lib/process-emails';
import { StatusPill } from './status-pill';

interface TaskItemProps {
    task: EmailTask;
}

export const TaskItem: React.FC<TaskItemProps> = ({ task }) => {
    return (
        <li className="flex justify-between items-center bg-gray-50 p-4 rounded-2xl border">
            <div className="flex-1 min-w-0">
                <p className="font-medium text-gray-800 truncate">{task.subject}</p>
                {task.error && <p className="text-xs text-red-500 truncate">Error: {task.error}</p>}
            </div>
            <StatusPill status={task.status} />
        </li>
    );
};