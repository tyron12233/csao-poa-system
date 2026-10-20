import React from 'react';

interface StatusPillProps {
    status: string;
}

export const STATUS_LABELS: Record<string, string> = {
    queued: 'Queued',
    fetching: 'Fetching',
    parsing: 'Reading',
    uploading_pdf: 'Uploading',
    building_request: 'Preparing',
    writing: 'Writing',
    done: 'Added',
    skipped: 'Skipped',
    held: 'Held',
    error: 'Error',
};

const STATUS_CLASSES: Record<string, string> = {
    done: 'bg-brand-50 text-brand-700 ring-brand-200',
    error: 'bg-red-50 text-red-700 ring-red-200',
    held: 'bg-amber-50 text-amber-800 ring-amber-200',
    skipped: 'bg-gray-100 text-gray-600 ring-gray-200',
    queued: 'bg-gray-50 text-gray-500 ring-gray-200',
};

export const StatusPill: React.FC<StatusPillProps> = ({ status }) => {
    const inProgress = !(status in STATUS_CLASSES);
    const className = STATUS_CLASSES[status] || 'bg-sky-50 text-sky-700 ring-sky-200';

    return (
        <span className={`inline-flex items-center gap-1.5 shrink-0 px-2 py-0.5 rounded-full text-[11px] font-semibold ring-1 ring-inset ${className}`}>
            {inProgress && <span className="h-1.5 w-1.5 rounded-full bg-current animate-pulse" />}
            {STATUS_LABELS[status] || status.replace('_', ' ')}
        </span>
    );
};
