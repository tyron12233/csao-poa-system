// components/ui/UserProfile.tsx
import React from 'react';

interface UserProfileProps {
    user: {
        name: string;
        picture: string;
    };
    onLogout: () => void;
}

export const UserProfile: React.FC<UserProfileProps> = ({ user, onLogout }) => {
    return (
        <div className="flex items-center space-x-3">
            <img src={user.picture} alt={user.name} className="h-10 w-10 rounded-full" />
            <button
                onClick={onLogout}
                className="text-sm bg-gray-200 hover:bg-gray-300 text-gray-800 font-semibold px-3 py-1 rounded-lg"
            >
                Logout
            </button>
        </div>
    );
};