import React, { useCallback } from 'react';
import { Mail, PenSquare } from 'lucide-react';
import { useUIStore } from '../../store/Store';
import { useAuth } from '../../auth/AuthProvider';
import { SearchInput } from './SearchInput';
import { UserMenu } from './UserMenu';

export const HeaderToolbar: React.FC = () => {
  const { setIsComposing, setComposeType, setComposeMailContext, searchQuery, setSearchQuery } = useUIStore();
  const auth = useAuth();

  const handleNewMail = useCallback(() => {
    setComposeType('new');
    setComposeMailContext(null);
    setIsComposing(true);
  }, [setComposeType, setComposeMailContext, setIsComposing]);

  return (
    <div className="h-14 min-h-[56px] border-b border-gray-200 flex items-center bg-gray-50 z-10 w-full gap-4 px-4">
      {/* Logo and New Mail Button */}
      <div className="flex items-center gap-2 text-blue-600">
        <Mail size={24} />
        <span className="font-semibold text-lg text-gray-800">WebMail</span>
        <button 
          onClick={handleNewMail}
          className="bg-blue-600 hover:bg-blue-700 text-white border-none rounded px-4 py-1.5 text-sm font-semibold cursor-pointer flex items-center gap-2 transition-colors ml-2"
        >
          <PenSquare size={16} />
          New Mail
        </button>
      </div>

      {/* Search Input - Outside Syncfusion to prevent focus loss */}
      <div className="flex-1 max-w-[400px]">
        <SearchInput value={searchQuery} onChange={setSearchQuery} />
      </div>

      {/* Icons - Using Syncfusion only for right-aligned icons */}
      <div className="flex gap-4 items-center text-gray-600 ml-auto">
        {/* <Bell size={20} className="cursor-pointer hover:text-blue-600 transition-colors" />
        <Settings size={20} className="cursor-pointer hover:text-blue-600 transition-colors" /> */}
        {auth.isAuthenticated ? (
          <UserMenu />
        ) : (
          <button
            onClick={() => auth.login()}
            className="bg-green-600 hover:bg-green-700 text-white rounded px-3 py-1 text-sm font-medium"
          >
            Login
          </button>
        )}
      </div>
    </div>
  );
};
