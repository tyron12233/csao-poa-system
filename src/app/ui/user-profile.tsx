import React from 'react';
import { Button } from './button';

interface UserProfileProps {
    user: {
        name: string;
        email?: string;
        picture: string;
    };
    onLogout: () => void;
}

export const UserProfile: React.FC<UserProfileProps> = ({ user, onLogout }) => {
    return (
        <div className="flex items-center gap-2">
            <img src={user.picture} alt="" className="h-8 w-8 rounded-full ring-1 ring-gray-200" referrerPolicy="no-referrer" />
            <div className="hidden sm:block text-right leading-tight min-w-0">
                <p className="text-sm font-medium text-gray-800 truncate max-w-[10rem]">{user.name}</p>
                {user.email && <p className="text-[11px] text-gray-500 truncate max-w-[10rem]">{user.email}</p>}
            </div>
            <Button variant="ghost" size="sm" className="whitespace-nowrap" onClick={onLogout}>Sign out</Button>
        </div>
    );
};
