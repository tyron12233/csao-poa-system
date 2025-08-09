import React from 'react';

interface StatusPillProps {
    status: string;
}

export const StatusPill: React.FC<StatusPillProps> = ({ status }) => {
    const statusClasses: { [key: string]: string } = {
        done: 'bg-green-100 text-green-800',
        error: 'bg-red-100 text-red-800',
        fetching: 'bg-blue-100 text-blue-800 animate-pulse',
        parsing: 'bg-yellow-100 text-yellow-800 animate-pulse',
        building_request: 'bg-indigo-100 text-indigo-800',
        writing: 'bg-purple-100 text-purple-800 animate-pulse',
        queued: 'bg-gray-200 text-gray-800',
        skipped: 'bg-gray-300 text-gray-600',
    };

    const className = statusClasses[status] || 'bg-gray-200';

    return (
        <span
            className={`ml-4 px-3 py-1 rounded-full text-xs font-semibold capitalize ${className}`}
        >
            {status.replace('_', ' ')}
        </span>
    );
};