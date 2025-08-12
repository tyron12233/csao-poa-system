
import React from 'react';
import type { EmailTask } from '../../lib/process-emails';
import { StatusPill } from './status-pill';

interface TaskItemProps {
    task: EmailTask;
    compact?: boolean;
    onClick?: () => void;
}

export const TaskItem: React.FC<TaskItemProps> = ({ task, compact = false, onClick }) => {
    if (compact) {
        return (
            <li
                title={task.subject}
                onClick={onClick}
                className={`group relative flex items-center gap-1 rounded-full px-3 py-1.5 border text-xs md:text-[11px] bg-white shadow-sm max-w-full transition-colors ${task.error ? 'border-red-300' : 'border-gray-200'} ${onClick ? 'cursor-pointer hover:border-blue-400 hover:shadow' : 'hover:border-gray-400'}`}
            >
                <span className="font-medium truncate max-w-[10rem] md:max-w-[14rem] text-gray-700">
                    {task.subject}
                </span>
                <StatusPill status={task.status} />
                {task.error && (
                    <span className="absolute left-1/2 -translate-x-1/2 top-full mt-1 whitespace-nowrap rounded bg-red-600 text-white px-2 py-1 text-[10px] shadow-lg opacity-0 group-hover:opacity-100 pointer-events-none">
                        {task.error}
                    </span>
                )}
            </li>
        );
    }
    return (
        <li onClick={onClick} className={`flex justify-between items-center bg-gray-50 p-4 rounded-2xl border ${onClick ? 'cursor-pointer hover:border-blue-400 hover:bg-blue-50' : ''}`}>
            <div className="flex-1 min-w-0">
                <p className="font-medium text-gray-800 truncate">{task.subject}</p>
                {task.error && <p className="text-xs text-red-500 truncate">Error: {task.error}</p>}
            </div>
            <StatusPill status={task.status} />
        </li>
    );
};