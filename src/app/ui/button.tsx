
import React from 'react';

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
    fullWidth?: boolean;
}

export const Button: React.FC<ButtonProps> = ({ children, fullWidth, ...props }) => {
    const widthClass = fullWidth ? 'w-full' : '';
    return (
        <button
            {...props}
            className={`bg-blue-600 hover:bg-blue-700 text-white font-bold px-6 py-3 rounded-2xl shadow-md transition-transform transform hover:scale-105 disabled:bg-gray-400 disabled:cursor-not-allowed disabled:transform-none ${widthClass}`}
        >
            {children}
        </button>
    );
};