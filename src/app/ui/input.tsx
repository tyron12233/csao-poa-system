import React from 'react';

export const Input: React.FC<React.InputHTMLAttributes<HTMLInputElement>> = ({ className = '', ...props }) => {
    return (
        <input
            {...props}
            className={`w-full px-3 py-2 text-sm text-gray-900 bg-white border border-gray-300 rounded-lg placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-transparent transition ${className}`}
        />
    );
};
