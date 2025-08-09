import React, { ReactNode, useEffect } from 'react';

interface ModalProps {
    open: boolean;
    title?: string;
    onClose: () => void;
    children: ReactNode;
    maxWidthClass?: string;
}

export const Modal: React.FC<ModalProps> = ({ open, title, onClose, children, maxWidthClass = 'max-w-lg' }) => {
    useEffect(() => {
        function onKey(e: KeyboardEvent) { if (e.key === 'Escape') onClose(); }
        if (open) document.addEventListener('keydown', onKey);
        return () => document.removeEventListener('keydown', onKey);
    }, [open, onClose]);

    if (!open) return null;
    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <div className="absolute inset-0 bg-black/40 backdrop-blur-sm" onClick={onClose} />
            <div className={`relative w-full ${maxWidthClass} bg-white rounded-2xl shadow-xl border border-gray-200 flex flex-col max-h-[80vh]`} role="dialog" aria-modal="true">
                <div className="flex items-start justify-between px-5 py-4 border-b border-gray-200">
                    <h2 className="text-base font-semibold text-gray-800">{title}</h2>
                    <button onClick={onClose} className="text-gray-400 hover:text-gray-600 focus:outline-none focus:ring-2 focus:ring-blue-400 rounded p-1" aria-label="Close dialog">✕</button>
                </div>
                <div className="overflow-y-auto px-5 py-4">
                    {children}
                </div>
            </div>
        </div>
    );
};
