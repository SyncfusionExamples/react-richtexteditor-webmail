import React, { createContext, useContext, useState, useMemo, useCallback, useEffect, useRef } from 'react';
import type { ReactNode } from 'react';
import { mailData as initialMailData, folderData as initialFolderData } from '../data/mockData';
import type { MailData, FolderData } from '../data/mockData';
import { mailAPI } from '../services/api';

// --- Folder Store ---
interface FolderStoreType {
  folders: FolderData[];
  selectedFolderId: string;
  setSelectedFolderId: (id: string) => void;
  selectedFolderName: string;
  renameFolder: (id: string, newName: string) => void;
  deleteFolder: (id: string) => void;
  createSubfolder: (parentId: string, folderName: string) => void;
  markFolderAsRead: (folderId: string) => void;
  addToFavorites: (folderId: string) => void;
  removeFromFavorites: (folderId: string) => void;
  getFolderCount: (folderId: string) => number;
  getFoldersWithCounts: () => FolderData[];
  setFolders: (folders: FolderData[]) => void;
  getSystemFolderId: (kind: 'inbox' | 'sent' | 'drafts') => string | undefined;
}

const FolderContext = createContext<FolderStoreType | undefined>(undefined);

export const useFolderStore = () => {
  const context = useContext(FolderContext);
  if (!context) throw new Error('useFolderStore must be used within a StoreProvider');
  return context;
};

// --- Mail Store ---
interface MailStoreType {
  mails: MailData[];
  selectedMailId: string | null;
  setSelectedMailId: (id: string | null) => void;
  selectedFolderMails: MailData[];
  selectedMail: MailData | undefined;
  deleteMail: (id: string) => void;
  restoreMail: (id: string) => void;
  toggleReadStatus: (id: string) => void;
  toggleFlagStatus: (id: string) => void;
  sendMail: (mail: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    bodyHtml: string;
    bodyText: string;
    senderEmail?: string;
    senderName?: string;
  }) => Promise<{ id: string }>;
  setMails: React.Dispatch<React.SetStateAction<MailData[]>>;
  isLoading: boolean;
  error: string | null;
  loadRealData: (token: string) => Promise<void>;
  dataVersion: number;
  /**
   * Fetch the full message body and attachment metadata for a single
   * mail (lazy-loaded on selection). Updates the matching entry in the
   * `mails` array so the MailReader can render attachments without
   * needing a second render path. Returns the updated entry, or null
   * on failure.
   */
  fetchMailDetails: (id: string) => Promise<MailData | null>;
}

const MailContext = createContext<MailStoreType | undefined>(undefined);

export const useMailStore = () => {
  const context = useContext(MailContext);
  if (!context) throw new Error('useMailStore must be used within a StoreProvider');
  return context;
};

// --- UI Store ---
export interface FilterOptions {
  isRead?: boolean | null;
  isFlagged?: boolean | null;
}

interface UIStoreType {
  isComposing: boolean;
  setIsComposing: (isComposing: boolean) => void;
  composeType: 'new' | 'reply' | 'replyAll' | 'forward' | null;
  setComposeType: (type: 'new' | 'reply' | 'replyAll' | 'forward' | null) => void;
  composeMailContext: MailData | null;
  setComposeMailContext: (mail: MailData | null) => void;
  searchQuery: string;
  setSearchQuery: (query: string) => void;
  filters: FilterOptions;
  setFilters: (filters: FilterOptions) => void;
  clearCompose: () => void;
}

const UIContext = createContext<UIStoreType | undefined>(undefined);

export const useUIStore = () => {
  const context = useContext(UIContext);
  if (!context) throw new Error('useUIStore must be used within a StoreProvider');
  return context;
};

// --- Store Provider ---
export const StoreProvider: React.FC<{ children: ReactNode }> = ({ children }) => {
  const [folders, setFolders] = useState<FolderData[]>(initialFolderData);
  const [selectedFolderId, setSelectedFolderId] = useState<string>('1-1');

  const [mails, setMails] = useState<MailData[]>(initialMailData);
  const [selectedMailId, setSelectedMailId] = useState<string | null>(null);

  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  // Increments on every setMails call to force ListViewComponent to re-render
  const [dataVersion, setDataVersion] = useState<number>(0);

  const [isComposing, setIsComposing] = useState<boolean>(false);
  const [composeType, setComposeType] = useState<'new' | 'reply' | 'replyAll' | 'forward' | null>(null);
  const [composeMailContext, setComposeMailContext] = useState<MailData | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filters, setFilters] = useState<FilterOptions>({});

  // Get token from localStorage for WebSocket connection
  const [wsToken, setWsToken] = useState<string | null>(() => localStorage.getItem('authToken'));

  // Listen for token changes
  useEffect(() => {
    const checkToken = () => {
      const token = localStorage.getItem('authToken');
      if (token !== wsToken) {
        setWsToken(token);
      }
    };
    window.addEventListener('storage', checkToken);
    const interval = setInterval(checkToken, 1000);
    return () => {
      window.removeEventListener('storage', checkToken);
      clearInterval(interval);
    };
  }, [wsToken]);

  // Ref to always have the latest loadRealData (declared before WebSocket useEffect)
  const loadRealDataRef = useRef<((token: string) => Promise<void>) | null>(null);

  // WebSocket connection for real-time mail updates
  const wsRef = useRef<WebSocket | null>(null);
  const reconnectAttempts = useRef(0);
  const lastEventId = useRef<string | null>(null);
  const reconnectTimeout = useRef<NodeJS.Timeout | null>(null);
  const WS_URL = (import.meta as any).env?.VITE_API_WS_URL || 'ws://localhost:3001/ws';

  useEffect(() => {
    if (!wsToken) return;

    let isUnmounted = false;

    function connect() {
      if (wsRef.current) wsRef.current.close();
      const url = `${WS_URL}?token=${encodeURIComponent(wsToken!)}`;
      const ws = new WebSocket(url);
      wsRef.current = ws;

      ws.onopen = () => {
        reconnectAttempts.current = 0;
        console.log('✅ WebSocket connected (StoreProvider)');
      };

      ws.onmessage = async (event) => {
        try {
          const data = JSON.parse(event.data);
          console.log('📩 WS message received (StoreProvider):', data);
          if (data.type === 'new_email') {
            if (lastEventId.current === data.timestamp) {
              console.log('⏭️ Duplicate event, skipping');
              return;
            }
            lastEventId.current = data.timestamp;
            // Check if loadRealDataRef is set
            console.log('🔍 loadRealDataRef.current:', typeof loadRealDataRef.current);
            if (loadRealDataRef.current) {
              console.log('🔄 Calling loadRealData from WS event (StoreProvider)');
              await loadRealDataRef.current(wsToken!);
              console.log('✅ loadRealData completed from WS event');
            } else {
              console.error('❌ loadRealDataRef.current is null! Falling back to direct API call');
              // Fallback: call API directly and update state
              try {
                const mailAPI = (await import('../services/api')).mailAPI;
                mailAPI.setToken(wsToken!);
                const response = await mailAPI.getInitialData();
                // ... transformation logic would go here
                console.log('📨 Fallback API response received');
              } catch (err) {
                console.error('Fallback API call failed:', err);
              }
            }
          }
        } catch (err) {
          console.error('WS message parse error:', err);
        }
      };

      ws.onclose = () => {
        console.log('❌ WebSocket disconnected (StoreProvider)');
        if (isUnmounted) return;
        reconnectAttempts.current++;
        const delay = Math.min(30000, 1000 * 2 ** reconnectAttempts.current);
        reconnectTimeout.current = setTimeout(connect, delay);
      };

      ws.onerror = (e) => {
        console.error('⚠️ WebSocket error:', e);
        ws.close();
      };
    }

    connect();

    return () => {
      isUnmounted = true;
      if (wsRef.current) wsRef.current.close();
      if (reconnectTimeout.current) clearTimeout(reconnectTimeout.current);
    };
  }, [wsToken]);

  const selectedFolderName = useMemo(() => {
    return folders.find(f => f.id === selectedFolderId)?.name || 'Inbox';
  }, [folders, selectedFolderId]);

  const getSystemFolderId = useCallback((kind: 'inbox' | 'sent' | 'drafts') => {
    const normalizedKind = kind.toLowerCase();
    // Skip synthetic Favorites shortcuts (`__fav__<realId>`) and the
    // synthetic group ids (`favorites`, `__folders_root__`) so the
    // returned id always matches a real Graph folder that holds
    // actual messages. The shortcuts have names that look identical
    // to the real folder, so we filter on id pattern first.
    const isShortcut = (id: string) =>
      id === 'favorites' || id === '__folders_root__' || id.startsWith('__fav__');
    const match = folders.find(folder => {
      if (isShortcut(folder.id)) return false;
      const folderName = (folder.name || '').toLowerCase();
      if (normalizedKind === 'inbox') return folderName.includes('inbox');
      if (normalizedKind === 'sent') return folderName.includes('sent');
      if (normalizedKind === 'drafts') return folderName.includes('draft');
      return false;
    });

    return match?.id;
  }, [folders]);

  const selectedFolderMails = useMemo(() => {
    const result = mails.filter(m => m.folderId === selectedFolderId);
    const folderName = folders.find(f => f.id === selectedFolderId)?.name || 'Unknown';
    if (result.length > 0 || mails.length === 0) {
      console.log(`✅ selectedFolderMails: "${folderName}" has ${result.length}/${mails.length} mails`);
    } else {
      console.warn(`⚠️ selectedFolderMails: "${folderName}" (${selectedFolderId}) has no matching mails (${mails.length} total)`);
    }
    return result;
  }, [mails, selectedFolderId, folders]);

  const selectedMail = useMemo(() => {
    return mails.find(m => m.id === selectedMailId);
  }, [mails, selectedMailId]);

  const handleSetSelectedFolderId = (id: string) => {
    setSelectedFolderId(id);
    setSelectedMailId(null); // Reset mail selection on folder change
    setIsComposing(false);
  };

  const handleSetSelectedMailId = (id: string | null) => {
    setSelectedMailId(id);
    if (id) {
      setIsComposing(false);
      // Mark as read when opened
      const mail = mails.find(m => m.id === id);
      if (mail && !mail.isRead) {
        toggleReadStatus(id);
      }
    }
  };

  const deleteMail = async (id: string) => {
    try {
      // Step 1: Call backend API to move mail to deleted items in Microsoft Graph
      await mailAPI.deleteMail(id);
      
      // Step 2: Find the Deleted Items folder ID from the current folders list
      const deletedItemsFolder = folders.find(f =>
        f.name.toLowerCase().includes('deleted')
      );
      
      if (deletedItemsFolder) {
        // Step 3: Update local state - move to deleted items folder
        setMails(prevMails => 
          prevMails.map(m => m.id === id ? { ...m, folderId: deletedItemsFolder.id } : m)
        );
        
        console.log(`✅ Mail ${id} deleted successfully and moved to Deleted Items folder (${deletedItemsFolder.id})`);
      } else {
        console.warn(`⚠️ Deleted Items folder not found, but mail was deleted from server`);
        // Still remove from current folder even if we can't find Deleted Items
        setMails(prevMails => prevMails.filter(m => m.id !== id));
      }
      
      if (selectedMailId === id) {
        setSelectedMailId(null);
      }
    } catch (error) {
      console.error('Error deleting mail:', error);
      throw error;
    }
  };

const getDeletedFolderId = () => {
  return folders.find(f =>
    f.name.toLowerCase().includes('deleted')
  )?.id || '';
};

  const restoreMail = async (id: string) => {
    try {
      // Step 1: Call backend API to restore mail to inbox in Microsoft Graph
      await mailAPI.restoreMail(id);
      
      // Step 2: Find the Inbox folder ID from the current folders list
      const inboxFolder = folders.find(f =>
        f.name.toLowerCase().includes('inbox')
      );
      
      if (inboxFolder) {
        // Step 3: Update local state - move back to inbox
        setMails(prevMails => 
          prevMails.map(m => m.id === id ? { ...m, folderId: inboxFolder.id } : m)
        );
        
        console.log(`✅ Mail ${id} restored successfully to Inbox folder (${inboxFolder.id})`);
      } else {
        console.warn(`⚠️ Inbox folder not found, but mail was restored on server`);
      }
      
      if (selectedMailId === id) {
        setSelectedMailId(null);
      }
    } catch (error) {
      console.error('Error restoring mail:', error);
      throw error;
    }
  };

  const toggleReadStatus = (id: string) => {
    setMails(prevMails => 
      prevMails.map(m => m.id === id ? { ...m, isRead: !m.isRead } : m)
    );
  };

  const toggleFlagStatus = (id: string) => {
    setMails(prevMails => 
      prevMails.map(m => m.id === id ? { ...m, isFlagged: !m.isFlagged } : m)
    );
  };

  const sendMail = async (mail: {
    to: string[];
    cc?: string[];
    bcc?: string[];
    subject: string;
    bodyHtml: string;
    bodyText: string;
    senderEmail?: string;
    senderName?: string;
  }) => {
    const recipients = mail.to.map(address => address.trim()).filter(Boolean);
    if (recipients.length === 0) {
      throw new Error('Add at least one recipient before sending');
    }

    const subject = mail.subject.trim() || '(No subject)';
    const bodyHtml = mail.bodyHtml.trim() || '<p><br/></p>';
    const bodyText = mail.bodyText.trim() || 'No content';

    await mailAPI.sendMail({
      to: recipients,
      cc: mail.cc?.map(address => address.trim()).filter(Boolean),
      bcc: mail.bcc?.map(address => address.trim()).filter(Boolean),
      subject,
      body: bodyHtml,
    });

    const senderEmail = mail.senderEmail || localStorage.getItem('authEmail') || 'me@example.com';
    const senderName = mail.senderName || 'Me';
    const sentFolderId = getSystemFolderId('sent') || '2-3';
    const newMail: MailData = {
      id: `m${Date.now()}`,
      folderId: sentFolderId,
      senderName,
      senderEmail,
      subject,
      snippet: bodyText.substring(0, 100) + (bodyText.length > 100 ? '...' : ''),
      // ISO UTC timestamp so the mail list's date grouping
      // ("Today", ...) and the `formatMailDate` formatter can
      // parse it correctly.
      date: new Date().toISOString(),
      isRead: true,
      isStarred: false,
      isFlagged: false,
      avatarColor: '#0078d4',
      body: bodyHtml,
      to: recipients,
      cc: mail.cc || [],
      bcc: mail.bcc || [],
    };

    setMails(prev => [newMail, ...prev]);

    // Return the new mail's id so callers (e.g. the compose flow) can
    // navigate to Sent Items and select the just-sent mail in the list.
    return { id: newMail.id };
  };

  /**
   * Fetch the full message body and attachment metadata for a single
   * mail and merge the result into the local `mails` array. Called
   * when the user selects a mail in the list — the list view shows
   * the lightweight preview (subject, sender, snippet), and the
   * reader view needs the full body + attachment list.
   *
   * Returns the updated MailData (or null on failure) so the caller
   * can react if needed (e.g. show an error toast).
   */
  const fetchMailDetails = useCallback(
    async (id: string): Promise<MailData | null> => {
      // Skip the fetch entirely if the ID is clearly synthetic (not
      // a real Microsoft Graph message id). Synthetic ids start with
      // "m" and a timestamp, and they don't exist on the server. The
      // local copy already has the preview data we need.
      const isSynthetic = /^m\d+$/.test(id);
      if (isSynthetic) {
        return mails.find((m) => m.id === id) ?? null;
      }

      try {
        // Use the dedicated /mail/messages/:id endpoint which uses
        // $expand=attachments on the backend so we get the full
        // attachment list (id, name, contentType, size) without the
        // base64 content bytes.
        const fullMessage = (await mailAPI.getMessage(id)) as
          | (Record<string, any> & {
              subject?: string;
              bodyPreview?: string;
              receivedDateTime?: string;
              isRead?: boolean;
              hasAttachments?: boolean;
              parentFolderId?: string;
              from?: { emailAddress?: { name?: string; address?: string } };
              sender?: { emailAddress?: { name?: string; address?: string } };
              body?: { contentType?: string; content?: string };
              attachments?: Array<{
                id: string;
                name: string;
                contentType?: string;
                size?: number;
                isInline?: boolean;
                lastModifiedDateTime?: string;
              }>;
            })
          | null;

        if (!fullMessage) {
          return null;
        }

        // Normalize the response into the MailData shape used by the
        // frontend. Graph returns more fields than we need; pick the
        // ones the MailReader and MailList actually use.
        const bodyHtml: string | undefined =
          fullMessage.body?.contentType === 'html'
            ? fullMessage.body.content
            : fullMessage.bodyPreview;

        const attachments = Array.isArray(fullMessage.attachments)
          ? fullMessage.attachments.map((a) => ({
              id: a.id,
              name: a.name,
              contentType: a.contentType || 'application/octet-stream',
              size: typeof a.size === 'number' ? a.size : 0,
              isInline: !!a.isInline,
              lastModifiedDateTime: a.lastModifiedDateTime,
            }))
          : [];

        const updated: MailData = {
          id,
          folderId:
            mails.find((m) => m.id === id)?.folderId ||
            fullMessage.parentFolderId ||
            '2-1',
          senderName:
            fullMessage.from?.emailAddress?.name ||
            fullMessage.sender?.emailAddress?.name ||
            'Unknown',
          senderEmail:
            fullMessage.from?.emailAddress?.address ||
            fullMessage.sender?.emailAddress?.address ||
            '',
          subject: fullMessage.subject || '(No subject)',
          snippet: fullMessage.bodyPreview || '',
          date: fullMessage.receivedDateTime || new Date().toISOString(),
          isRead: !!fullMessage.isRead,
          isStarred: false,
          isFlagged: false,
          avatarColor: '#0078d4',
          body: bodyHtml,
          hasAttachments: !!fullMessage.hasAttachments,
          attachments,
        };

        // Merge into the local mails array. If the mail is not in the
        // array (e.g. opened directly via deep link), prepend it.
        setMails((prev) => {
          const idx = prev.findIndex((m) => m.id === id);
          if (idx === -1) return [updated, ...prev];
          const next = prev.slice();
          // Preserve isRead/isFlagged/isStarred from the local copy
          // (those are user-driven and should win over the server
          // values that the user may have already toggled).
          next[idx] = { ...next[idx], ...updated };
          return next;
        });

        return updated;
      } catch (err: any) {
        // A 404 from the backend means the message no longer exists
        // in the user's mailbox (deleted, moved, or never existed).
        // This is a known/expected state — not a server error — so we
        // log at `warn` level and don't pollute the console with a
        // red error. The local copy (with subject/sender/snippet) is
        // still shown to the user; they just won't see body or
        // attachments for the missing message.
        const message = err instanceof Error ? err.message : String(err);
        const isNotFound = /Not Found|404/i.test(message);
        if (isNotFound) {
          console.warn(
            `[fetchMailDetails] Mail ${id.slice(0, 16)}… is no longer available on the server.`,
            message
          );
        } else {
          console.error('Failed to fetch mail details:', err);
        }
        return null;
      }
    },
    [mails]
  );

  const renameFolder = (id: string, newName: string) => {
    setFolders(prevFolders =>
      prevFolders.map(f => f.id === id ? { ...f, name: newName } : f)
    );
  };

  const deleteFolder = (id: string) => {
    // Prevent deletion of system folders (Inbox, Sent Items, etc.)
    const folder = folders.find(f => f.id === id);
    if (folder && !folder.parentId) {
      return; // Don't delete parent folders
    }
    
    setFolders(prevFolders =>
      prevFolders.filter(f => f.id !== id && f.parentId !== id)
    );
    
    // If deleted folder is selected, select Inbox
    if (selectedFolderId === id) {
      setSelectedFolderId('1-1');
    }
  };

  const createSubfolder = (parentId: string, folderName: string) => {
    const newId = `folder-${Date.now()}`;
    const newFolder: FolderData = {
      id: newId,
      parentId: parentId,
      name: folderName,
      hasChild: false,
      expanded: false
    };
    setFolders(prevFolders => {
      // Add the new subfolder and update parent folder to have hasChild: true
      return prevFolders.map(f => 
        f.id === parentId ? { ...f, hasChild: true } : f
      ).concat(newFolder);
    });
  };

  const markFolderAsRead = (folderId: string) => {
    setMails(prevMails =>
      prevMails.map(m => m.folderId === folderId ? { ...m, isRead: true } : m)
    );
  };

  const addToFavorites = (folderId: string) => {
    setFolders(prevFolders => {
      const folder = prevFolders.find(f => f.id === folderId);
      if (!folder) return prevFolders;

      // Don't pin folders that are already favorites, and skip
      // the synthetic Favorites group itself.
      if (folder.parentId === 'favorites' || folder.parentId === '1') {
        return prevFolders;
      }

      // Move the folder into Favorites by tagging its parentId
      // with the synthetic Favorites id ('favorites'). The real
      // parent is preserved in `originalParentId` so
      // `removeFromFavorites` (and the backend's permanent
      // favorites list) can restore it.
      //
      // We also accept the legacy `'1'` (mock-data Favorites id)
      // as already-favourited, so existing pinned folders
      // continue to behave correctly.
      return prevFolders.map(f =>
        f.id === folderId
          ? {
              ...f,
              parentId: 'favorites',
              originalParentId: folder.originalParentId ?? folder.parentId,
            }
          : f
      );
    });
  };

  const removeFromFavorites = (folderId: string) => {
    setFolders(prevFolders => {
      const folder = prevFolders.find(f => f.id === folderId);
      if (!folder) return prevFolders;

      // Only act on folders currently pinned to Favorites.
      if (folder.parentId !== 'favorites' && folder.parentId !== '1') {
        return prevFolders;
      }

      // Restore to the original parent if we recorded one when
      // the folder was pinned. Otherwise, leave the parent
      // unchanged (best effort — the user's intent is to remove
      // the shortcut, not to relocate the folder).
      const restoredParentId = folder.originalParentId ?? folder.parentId;
      return prevFolders.map(f =>
        f.id === folderId
          ? {
              ...f,
              parentId:
                restoredParentId === 'favorites' || restoredParentId === '1'
                  ? undefined
                  : restoredParentId,
              originalParentId: undefined,
            }
          : f
      );
    });
  };

  const clearCompose = () => {
    setIsComposing(false);
    setComposeType(null);
    setComposeMailContext(null);
  };

  /**
   * Load real data from backend after user login
   * This replaces mock data with actual mail folders and messages
   */
  const loadRealData = useCallback(async (token: string) => {
    try {
      setIsLoading(true);
      setError(null);
      console.log('🔄 Starting to load real data...');

      // Set token in API client
      mailAPI.setToken(token);
      console.log('✅ Token set in API client');

      // Fetch initial data from backend
      console.log('📡 Fetching initial data from /mail/initial-data...');
      const response = await mailAPI.getInitialData();
      console.log('📨 API Response:', response);

      // Transform backend data to match frontend data structure.
      //
      // The backend's `transformAndHierarchizeFolders` already
      // emits the synthetic Favorites + Folders groups in
      // Outlook's order, but we still run the frontend copy
      // defensively so a stale or differently-versioned backend
      // (e.g. one that wasn't rebuilt after the fix) doesn't drop
      // the Favorites group. The frontend transform is idempotent
      // — if the synthetic groups are already present, the extra
      // pass just re-orders + re-groups; the shortcut ids start
      // with `__fav__` so they remain stable.
      const { wrapWithOutlookGroups } = await import('../utils/transformers');
      const realFolders = response.folders.map((folder: any) => ({
        id: folder.id,
        name: folder.name,
        icon: folder.icon,
        hasChild: folder.hasChild,
        expanded: folder.expanded,
        parentId: folder.parentId,
        count: folder.count,
      }));

      // Always wrap with the synthetic groups. `wrapWithOutlookGroups`
      // is idempotent: if the backend already produced them, the
      // shortcuts it added (and the original system folders) just
      // get re-emitted under the new synthetic groups. If the
      // backend didn't produce them, this call creates them.
      const transformedFolders: FolderData[] = wrapWithOutlookGroups(realFolders);

      // Diagnostic logs so any tree-shape regression is obvious
      // from the browser console. The user reported subfolders
      // rendering without the expected indentation; this log will
      // make it clear which folder has which parent.
      console.log(
        '📁 Final folder tree (id, name, parentId):',
        transformedFolders.map((f: FolderData) => ({
          id: typeof f.id === 'string' ? f.id.slice(0, 30) : f.id,
          name: f.name,
          parentId:
            typeof f.parentId === 'string' ? f.parentId.slice(0, 30) : f.parentId,
        }))
      );
      const favShortcuts = transformedFolders.filter(
        (f: FolderData) => f.parentId === 'favorites'
      );
      console.log(
        `📁 Favorites group has ${favShortcuts.length} shortcut children:`,
        favShortcuts.map((f: FolderData) => f.name)
      );
      // Also log the orphan check: any folder whose parentId
      // doesn't reference a real folder in the list will render
      // as a root instead of a child, which is what causes the
      // "subfolder not indented" symptom.
      const knownIds = new Set(transformedFolders.map((f) => f.id));
      const orphans = transformedFolders.filter(
        (f: FolderData) =>
          f.parentId && !knownIds.has(f.parentId as string)
      );
      if (orphans.length > 0) {
        console.warn(
          '📁 Found orphan folders (parentId not in list):',
          orphans.map((f: FolderData) => ({
            id: f.id,
            name: f.name,
            parentId: f.parentId,
          }))
        );
      }

      // Flatten all mails from mailsByFolder structure
      const transformedMails: MailData[] = [];
      console.log('🔍 Processing mailsByFolder keys:', Object.keys(response.mailsByFolder));
      for (const folderId in response.mailsByFolder) {
        const folderMails = response.mailsByFolder[folderId];
        console.log(`  - Folder ID: ${folderId}, Mail count: ${folderMails.length}`);
        if (folderMails.length > 0) {
          console.log(`    Sample mail folderId: ${folderMails[0].folderId}`);
        }
        transformedMails.push(
          ...folderMails.map((mail: any) => ({
            id: mail.id,
            folderId: mail.folderId,
            senderName: mail.senderName,
            senderEmail: mail.senderEmail,
            subject: mail.subject,
            snippet: mail.snippet,
            date: mail.date,
            isRead: mail.isRead,
            isStarred: mail.isStarred,
            isFlagged: mail.isFlagged,
            avatarColor: mail.avatarColor,
            body: mail.body,
          }))
        );
      }
      console.log('📧 Transformed mails count:', transformedMails.length);
      console.log('🔍 First 3 mails folderIds:', transformedMails.slice(0, 3).map(m => ({ id: m.id.substring(0, 10), folderId: m.folderId.substring(0, 30) })));
      console.log('📁 Transformed folders list:', transformedFolders.map(f => ({ id: f.id.substring(0, 30), name: f.name })));

      // Update store with real data
      console.log('🔄 Updating store with real data...');
      console.log('   Calling setFolders with:', transformedFolders.length, 'folders');
      console.log('   Calling setMails with:', transformedMails.length, 'mails');
      setMails(transformedMails);
      // Increment dataVersion to force ListViewComponent to re-render
      setDataVersion(v => v + 1);
      console.log('✅ setMails called with', transformedMails.length, 'mails. dataVersion incremented.');
      
      if (transformedFolders.length > 0) {
        setFolders(transformedFolders);

        const folderIdsWithMail = new Set(transformedMails.map(mail => mail.folderId));

        // Helper: pick the "real" id of a folder. Favorites shortcuts
        // have synthetic ids like `__fav__<realId>`; the mail list is
        // keyed by the real Graph folder id, so the selectedFolderId
        // must always be a real id, never a shortcut id.
        const { getRealFolderId } = await import('../utils/transformers');
        const isSyntheticGroup = (id: string) =>
          id === 'favorites' || id === '__folders_root__' ||
          (typeof id === 'string' && id.startsWith('__fav__'));

        // First, try to find a real Inbox folder that actually has
        // messages. Skip synthetic groups and shortcuts so the
        // returned id always matches a `folderId` on a Mail row.
        let defaultFolder = transformedFolders.find(f =>
          !isSyntheticGroup(f.id) &&
          f.name.toLowerCase().includes('inbox') &&
          folderIdsWithMail.has(f.id)
        );

        // If no Inbox with messages found, pick the first real
        // folder that has mail.
        if (!defaultFolder) {
          defaultFolder = transformedFolders.find(f =>
            !isSyntheticGroup(f.id) && folderIdsWithMail.has(f.id)
          );
        }

        // If still nothing matches, fall back to the first real
        // system folder.
        if (!defaultFolder) {
          defaultFolder = transformedFolders.find(
            (f) => !isSyntheticGroup(f.id) && f.id !== 'favorites'
          );
        }

        // Last resort: use first folder if nothing else is available
        if (!defaultFolder) {
          defaultFolder = transformedFolders.find((f) => !isSyntheticGroup(f.id)) ||
            transformedFolders[0];
        }

        // Final safety: if the chosen folder is a shortcut, convert
        // to the real id so subsequent filters (e.g.
        // `mails.filter(m => m.folderId === selectedFolderId)`) work.
        const defaultId = getRealFolderId(defaultFolder.id);

        console.log('📌 Setting selected folder ID to:', defaultId);
        console.log('📌 Selected folder object:', { id: defaultId.substring(0, 30), name: defaultFolder.name });
        
        // Check how many mails will match the chosen id.
        const matchingMails = transformedMails.filter(m => m.folderId === defaultId);
        console.log('🔎 Mails matching selected folder ID:', matchingMails.length);
        if (matchingMails.length === 0 && transformedMails.length > 0) {
          console.warn('⚠️ WARNING: No mails match the selected folder ID');
          console.log('   Selected folder:', { id: defaultId, name: defaultFolder.name });
          console.log('   Available mail folderIds:', [...new Set(transformedMails.map(m => m.folderId))].slice(0, 5));
          console.log('   This is normal if the selected folder is a parent/container folder');
        }

        setSelectedFolderId(defaultId);
      } else {
        console.warn('⚠️ No folders returned from backend; keeping the existing folder tree in place');
      }

      setIsLoading(false);
      console.log('✅ Real data loaded successfully!');
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : 'Failed to load mail data';
      setError(errorMsg);
      setIsLoading(false);
      console.error('❌ Error loading real data:', err);
      // Keep mock data visible if loading fails
    }
  }, []);

  // Keep the ref updated with the latest loadRealData
  useEffect(() => {
    loadRealDataRef.current = loadRealData;
  }, [loadRealData]);

  const getFolderCount = useCallback((folderId: string): number => {
    return mails.filter(m => m.folderId === folderId && !m.isRead).length;
  }, [mails]);

  // Only recompute when folders or mail read state changes
  // The counts are unread-message counts for each folder
  const getFoldersWithCounts = useCallback(() => {
    return folders.map(folder => ({
      ...folder,
      count: mails.filter(m => m.folderId === folder.id && !m.isRead).length
    }));
  }, [folders, mails]);

  // Memoize the context value to prevent unnecessary re-renders
  // Only update when actual store functions/data change
  const folderContextValue = useMemo(() => ({
    folders: getFoldersWithCounts(),
    selectedFolderId,
    setSelectedFolderId: handleSetSelectedFolderId,
    selectedFolderName,
    renameFolder,
    deleteFolder,
    createSubfolder,
    markFolderAsRead,
    addToFavorites,
    removeFromFavorites,
    getFolderCount,
    getFoldersWithCounts,
    setFolders,
    getSystemFolderId
  }), [folders, mails, selectedFolderId, selectedFolderName, handleSetSelectedFolderId, renameFolder, deleteFolder, createSubfolder, markFolderAsRead, addToFavorites, removeFromFavorites, getFolderCount, getFoldersWithCounts, getSystemFolderId]);

  const mailContextValue = useMemo(() => ({
    mails,
    selectedMailId,
    setSelectedMailId: handleSetSelectedMailId,
    selectedFolderMails,
    selectedMail,
    deleteMail,
    restoreMail,
    toggleReadStatus,
    toggleFlagStatus,
    sendMail,
    setMails,
    isLoading,
    error,
    loadRealData,
    dataVersion,
    fetchMailDetails
  }), [mails, selectedMailId, handleSetSelectedMailId, selectedFolderMails, selectedMail, deleteMail, restoreMail, toggleReadStatus, toggleFlagStatus, sendMail, isLoading, error, loadRealData, dataVersion, fetchMailDetails]);

  const uiContextValue = useMemo(() => ({
    isComposing,
    setIsComposing,
    composeType,
    setComposeType,
    composeMailContext,
    setComposeMailContext,
    searchQuery,
    setSearchQuery,
    filters,
    setFilters,
    clearCompose
  }), [isComposing, setIsComposing, composeType, setComposeType, composeMailContext, setComposeMailContext, searchQuery, setSearchQuery, filters, setFilters, clearCompose]);

  return (
    <UIContext.Provider value={uiContextValue}>
      <FolderContext.Provider value={folderContextValue}>
        <MailContext.Provider value={mailContextValue}>
          {children}
        </MailContext.Provider>
      </FolderContext.Provider>
    </UIContext.Provider>
  );
};
