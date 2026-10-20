import React from 'react';

type ButtonVariant = 'primary' | 'secondary' | 'ghost';
type ButtonSize = 'sm' | 'md' | 'lg';

interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
    fullWidth?: boolean;
    variant?: ButtonVariant;
    size?: ButtonSize;
}

const VARIANTS: Record<ButtonVariant, string> = {
    primary: 'bg-brand-600 text-white shadow-sm hover:bg-brand-700 disabled:bg-gray-300 disabled:text-gray-500 disabled:shadow-none',
    secondary: 'bg-white text-gray-800 border border-gray-300 shadow-sm hover:bg-gray-50 hover:border-gray-400 disabled:text-gray-400 disabled:bg-gray-50',
    ghost: 'text-gray-600 hover:text-gray-900 hover:bg-gray-100 disabled:text-gray-300',
};

const SIZES: Record<ButtonSize, string> = {
    sm: 'text-xs px-2.5 py-1.5 rounded-lg',
    md: 'text-sm px-4 py-2 rounded-lg',
    lg: 'text-base px-5 py-3 rounded-xl',
};

export const Button: React.FC<ButtonProps> = ({ children, fullWidth, variant = 'primary', size = 'md', className = '', type = 'button', ...props }) => {
    return (
        <button
            type={type}
            {...props}
            className={`inline-flex items-center justify-center gap-2 font-semibold transition-colors focus:outline-none focus-visible:ring-2 focus-visible:ring-brand-500 focus-visible:ring-offset-2 disabled:cursor-not-allowed ${VARIANTS[variant]} ${SIZES[size]} ${fullWidth ? 'w-full' : ''} ${className}`}
        >
            {children}
        </button>
    );
};
