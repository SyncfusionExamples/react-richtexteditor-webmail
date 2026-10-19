import React, { createContext, useContext, useEffect, useState, useCallback } from 'react';
import type { ReactNode } from 'react';
import { useMailStore } from '../store/Store';
import { mailAPI } from '../services/api';

const DEFAULT_API = import.meta.env.VITE_API_URL || 'http://localhost:3001';

export interface UserProfile {
  id: string;
  email: string;
  displayName: string;
  jobTitle?: string | null;
  hasPhoto: boolean;
}

interface AuthContextType {
  token: string | null;
  email: string | null;
  displayName: string | null;
  profile: UserProfile | null;
  isAuthenticated: boolean;
  login: () => Promise<void>;
  logout: () => void;
  refreshProfile: () => Promise<void>;
  isLoadingData: boolean;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const useAuth = () => {
  const ctx = useContext(AuthContext);
  if (!ctx) throw new Error('useAuth must be used within AuthProvider');
  return ctx;
};

export const AuthProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const queryParams = new URLSearchParams(window.location.search);
  const initialToken = queryParams.get('token');
  const initialEmail = queryParams.get('email');
  const initialName = queryParams.get('name');

  const [token, setToken] = useState<string | null>(() => localStorage.getItem('authToken') || initialToken);
  const [email, setEmail] = useState<string | null>(() => {
    const savedEmail = localStorage.getItem('authEmail');
    return savedEmail || (initialEmail ? decodeURIComponent(initialEmail) : null);
  });
  const [displayName, setDisplayName] = useState<string | null>(() => {
    const saved = localStorage.getItem('authDisplayName');
    if (saved) return saved;
    if (initialName) return decodeURIComponent(initialName);
    return null;
  });
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [isLoadingData, setIsLoadingData] = useState<boolean>(() => !!initialToken);

  // Get mail and folder stores to load real data after login
  const { loadRealData } = useMailStore();

  useEffect(() => {
    // If redirect contains token, email, and/or name, capture and store
    const params = new URLSearchParams(window.location.search);
    const t = params.get('token');
    const e = params.get('email');
    const n = params.get('name');

    if (t) {
      localStorage.setItem('authToken', t);
      setToken(t);
    }
    if (e) {
      const decoded = decodeURIComponent(e);
      localStorage.setItem('authEmail', decoded);
      setEmail(decoded);
    }
    if (n) {
      const decodedName = decodeURIComponent(n);
      localStorage.setItem('authDisplayName', decodedName);
      setDisplayName(decodedName);
    }

    // Clean URL to remove sensitive query params
    if (t || e || n) {
      const newUrl = window.location.origin + window.location.pathname;
      window.history.replaceState({}, document.title, newUrl);
    }
  }, []);

  // Pull the freshest profile (displayName, hasPhoto) from the backend
  // whenever we have a token. This is what powers the avatar/name in
  // the header toolbar.
  const refreshProfile = useCallback(async () => {
    if (!token) return;
    try {
      const fresh = await mailAPI.getMe();
      setProfile(fresh);
      if (fresh.displayName) {
        setDisplayName(fresh.displayName);
        localStorage.setItem('authDisplayName', fresh.displayName);
      }
    } catch (err) {
      console.warn('Failed to refresh user profile:', err);
    }
  }, [token]);

  useEffect(() => {
    if (!token) return;

    mailAPI.setToken(token);
    setIsLoadingData(true);

    loadRealData(token)
      .catch(err => console.error('Failed to load mail data:', err))
      .finally(() => setIsLoadingData(false));

    refreshProfile();
  }, [token, loadRealData, refreshProfile]);

  const login = async () => {
    const api = DEFAULT_API;
    const res = await fetch(`${api}/auth/login`);
    if (!res.ok) throw new Error('Failed to get auth url');
    const data = await res.json();
    const authUrl = data.authUrl || data.url || data.redirectUrl;
    if (!authUrl) throw new Error('Auth URL not returned from server');
    window.location.href = authUrl;
  };

  const logout = () => {
    localStorage.removeItem('authToken');
    localStorage.removeItem('authEmail');
    localStorage.removeItem('authDisplayName');
    setToken(null);
    setEmail(null);
    setDisplayName(null);
    setProfile(null);
    // Optionally reload to reset app state
    window.location.href = '/';
  };

  return (
    <AuthContext.Provider
      value={{
        token,
        email,
        displayName,
        profile,
        isAuthenticated: !!token,
        login,
        logout,
        refreshProfile,
        isLoadingData,
      }}
    >
      {children}
    </AuthContext.Provider>
  );
};

export default AuthProvider;
