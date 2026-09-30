import React, { useEffect, useRef, useState } from 'react';
import { LogOut, RefreshCw } from 'lucide-react';
import { useAuth } from '../../auth/AuthProvider';
import { mailAPI } from '../../services/api';

/**
 * Small palette used for the initials-circle fallback when the user
 * has no Microsoft Graph profile photo. Re-uses the same brand colors
 * already used for sender avatars in the mail list so the UI feels
 * consistent.
 */
const FALLBACK_PALETTE = [
  '#0078d4', '#107c10', '#5c2d91', '#e3008c',
  '#d13438', '#ca5010', '#038387', '#8764b8',
];

const pickColor = (seed: string): string => {
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = (hash * 31 + seed.charCodeAt(i)) | 0;
  }
  return FALLBACK_PALETTE[Math.abs(hash) % FALLBACK_PALETTE.length];
};

const getInitials = (name: string, email: string): string => {
  const source = (name && name.trim()) || email || '';
  if (!source) return '?';
  const parts = source.split(/[\s@.]+/).filter(Boolean);
  if (parts.length === 0) return source.charAt(0).toUpperCase();
  if (parts.length === 1) return parts[0].charAt(0).toUpperCase();
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
};

export const UserMenu: React.FC = () => {
  const auth = useAuth();
  const [photoUrl, setPhotoUrl] = useState<string | null>(null);
  const [photoFailed, setPhotoFailed] = useState(false);
  const [open, setOpen] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const wrapperRef = useRef<HTMLDivElement | null>(null);
  const lastPhotoKeyRef = useRef<string | null>(null);

  const displayName = auth.displayName || auth.email || 'User';
  const fallbackColor = pickColor(auth.email || displayName);
  const initials = getInitials(displayName, auth.email || '');

  // Pull the Graph photo as an object URL once we know the user is
  // authenticated. We only re-fetch when the user identity changes,
  // never on every render.
  useEffect(() => {
    if (!auth.isAuthenticated) {
      setPhotoUrl(null);
      setPhotoFailed(false);
      lastPhotoKeyRef.current = null;
      return;
    }

    let cancelled = false;
    let createdUrl: string | null = null;

    const loadPhoto = async () => {
      try {
        const result = await mailAPI.getMyPhoto();
        if (cancelled) return;
        if (result) {
          createdUrl = URL.createObjectURL(result.blob);
          setPhotoUrl(createdUrl);
          setPhotoFailed(false);
        } else {
          setPhotoUrl(null);
        }
      } catch (err) {
        if (!cancelled) {
          console.warn('Could not load user photo:', err);
          setPhotoUrl(null);
        }
      }
    };

    const photoKey = auth.profile?.id || auth.email || 'me';
    if (lastPhotoKeyRef.current !== photoKey) {
      lastPhotoKeyRef.current = photoKey;
      loadPhoto();
    }

    return () => {
      cancelled = true;
      if (createdUrl) URL.revokeObjectURL(createdUrl);
    };
  }, [auth.isAuthenticated, auth.profile?.id, auth.email]);

  // Close the dropdown when clicking outside of it.
  useEffect(() => {
    if (!open) return;
    const handleClick = (e: MouseEvent) => {
      if (wrapperRef.current && !wrapperRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClick);
    return () => document.removeEventListener('mousedown', handleClick);
  }, [open]);

  const handleRefresh = async () => {
    setRefreshing(true);
    try {
      await auth.refreshProfile();
    } finally {
      setRefreshing(false);
      setOpen(false);
    }
  };

  return (
    <div className="relative" ref={wrapperRef}>
      <button
        type="button"
        onClick={() => setOpen(prev => !prev)}
        className="flex items-center gap-2 rounded-full pl-1 pr-3 py-1 hover:bg-gray-200 transition-colors"
        title={auth.email || ''}
      >
        {photoUrl && !photoFailed ? (
          <img
            src={photoUrl}
            alt={displayName}
            className="w-8 h-8 rounded-full object-cover border border-gray-200"
            onError={() => setPhotoFailed(true)}
          />
        ) : (
          <span
            className="w-8 h-8 rounded-full flex items-center justify-center text-white text-sm font-semibold border border-gray-200"
            style={{ backgroundColor: fallbackColor }}
            aria-hidden="true"
          >
            {initials}
          </span>
        )}
        <span className="text-sm font-medium text-gray-800 hidden sm:inline max-w-[160px] truncate">
          {displayName}
        </span>
      </button>

      {open && (
        <div
          className="absolute right-0 mt-2 w-64 bg-white border border-gray-200 rounded-lg shadow-lg z-50 overflow-hidden"
          role="menu"
        >
          <div className="px-4 py-3 border-b border-gray-100 flex items-center gap-3">
            {photoUrl && !photoFailed ? (
              <img
                src={photoUrl}
                alt={displayName}
                className="w-10 h-10 rounded-full object-cover"
              />
            ) : (
              <span
                className="w-10 h-10 rounded-full flex items-center justify-center text-white text-sm font-semibold"
                style={{ backgroundColor: fallbackColor }}
              >
                {initials}
              </span>
            )}
            <div className="min-w-0">
              <div className="text-sm font-semibold text-gray-900 truncate">
                {displayName}
              </div>
              <div className="text-xs text-gray-500 truncate" title={auth.email || ''}>
                {auth.email}
              </div>
              {auth.profile?.jobTitle && (
                <div className="text-xs text-gray-500 truncate">
                  {auth.profile.jobTitle}
                </div>
              )}
            </div>
          </div>

          <button
            type="button"
            onClick={handleRefresh}
            disabled={refreshing}
            className="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center gap-2 disabled:opacity-60"
            role="menuitem"
          >
            <RefreshCw size={14} className={refreshing ? 'animate-spin' : ''} />
            {refreshing ? 'Refreshing…' : 'Refresh profile'}
          </button>

          <button
            type="button"
            onClick={() => {
              setOpen(false);
              auth.logout();
            }}
            className="w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-red-50 flex items-center gap-2 border-t border-gray-100"
            role="menuitem"
          >
            <LogOut size={14} />
            Logout
          </button>
        </div>
      )}
    </div>
  );
};

export default UserMenu;
