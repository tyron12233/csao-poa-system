import React from 'react';

interface CardProps {
    children: React.ReactNode;
}

export const Card: React.FC<CardProps> = ({ children }) => {
    return (
        <div className="w-full max-w-2xl bg-white rounded-2xl border border-gray-200 shadow-sm p-5 sm:p-8">
            {children}
        </div>
    );
};
