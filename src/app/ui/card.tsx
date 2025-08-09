
import React from 'react';

interface CardProps {
    children: React.ReactNode;
}

export const Card: React.FC<CardProps> = ({ children }) => {
    return (
        <div className="w-full max-w-2xl bg-white rounded-3xl shadow-xl p-8">
            {children}
        </div>
    );
};