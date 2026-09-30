/**
 * Transform Microsoft Graph API responses to match frontend expectations
 */

// Color palette for avatar generation
const COLORS = [
  '#0078d4', '#107c10', '#d13438', '#4c4a48', '#ffc646', '#7fba00', '#7c4a7e', '#f7630c'
];

// Generate a consistent color for a given string (sender email/name)
function generateColorFromString(str: string): string {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash) + str.charCodeAt(i);
    hash = hash & hash; // Convert to 32bit integer
  }
  return COLORS[Math.abs(hash) % COLORS.length];
}

/**
 * Transform Microsoft Graph Folder to Frontend FolderData
 */
export interface FrontendFolder {
  id: string;
  name: string;
  icon?: string;
  hasChild?: boolean;
  expanded?: boolean;
  parentId?: string;
  count?: number; // unread count
}

export interface GraphFolder {
  id: string;
  displayName: string;
  parentFolderId?: string;
  childFolderCount?: number;
  unreadItemCount?: number;
  /**
   * Graph-assigned well-known folder identifier. We do NOT request
   * this field from the API (some tenants reject it in a flat
   * `select` with: "Could not find a property named
   * 'wellKnownName' on type 'microsoft.graph.mailFolder'."), so it
   * will usually be `undefined` in practice. The canonical Outlook
   * ordering is resolved from `displayName` instead, with
   * `wellKnownName` used as a tie-breaker when present.
   */
  wellKnownName?: string;
  totalItemCount?: number;
}

export function transformGraphFolderToFrontend(
  graphFolder: GraphFolder,
  unreadCount: number = 0
): FrontendFolder {
  // Map Graph folder names to icons
  const folderNameLower = graphFolder.displayName.toLowerCase();
  let icon: string | undefined;

  if (folderNameLower.includes('inbox')) icon = 'Inbox';
  else if (folderNameLower.includes('sent') || folderNameLower.includes('outbox')) icon = 'Send';
  else if (folderNameLower.includes('draft')) icon = 'FileText';
  else if (folderNameLower.includes('delete') || folderNameLower.includes('trash')) icon = 'Trash2';
  else if (folderNameLower.includes('archive')) icon = 'Archive';
  else if (folderNameLower.includes('junk') || folderNameLower.includes('spam')) icon = 'AlertCircle';

  return {
    id: graphFolder.id,
    name: graphFolder.displayName,
    icon,
    hasChild: (graphFolder.childFolderCount || 0) > 0,
    expanded: folderNameLower.includes('inbox'), // Expand Inbox by default
    parentId: graphFolder.parentFolderId,
    count: unreadCount || graphFolder.unreadItemCount || 0,
  };
}

/**
 * Canonical order in which Outlook renders its well-known folders.
 * Keys match Microsoft Graph's `wellKnownName` values exactly.
 * Index = priority (lower = higher in the list). Custom user
 * folders aren't listed here — they sort after all well-known
 * folders in the order Graph returns them (which matches the
 * order Outlook shows them in the folder pane).
 */
const WELL_KNOWN_ORDER: Record<string, number> = {
  inbox: 0,
  drafts: 1,
  sentitems: 2,
  deleteditems: 3,
  junkemail: 4,
  archive: 5,
  notes: 6,
  outbox: 7,
};

/**
 * Resolve the canonical Outlook rank for a folder. Checks
 * `wellKnownName` first (when present), then falls back to a
 * display-name match so the UI still gets Outlook-style ordering
 * when Graph omits `wellKnownName` (older tenants, custom
 * renames, etc.).
 *
 * Top-level system folders have predictable English labels:
 * "Inbox", "Drafts", "Sent Items", "Deleted Items", "Junk
 * Email", "Archive", "Notes", "Outbox". We strip a trailing
 * parenthesized suffix like "Sent Items (Archive)" and compare
 * the whole core, so user folders that happen to contain a
 * generic word (e.g. "Inbox Receipts", "Project Archive") are
 * NOT misclassified as system folders.
 */
function wellKnownRank(folder: GraphFolder): number {
  const wk = (folder.wellKnownName || '').toLowerCase();
  if (wk && WELL_KNOWN_ORDER[wk] !== undefined) {
    return WELL_KNOWN_ORDER[wk];
  }

  const name = (folder.displayName || '').toLowerCase().trim();
  // Strip a trailing parenthesized suffix, e.g. "Sent Items (Archive)".
  const core = name.replace(/\s*\([^)]*\)\s*$/, '').trim();

  if (core === 'inbox') return 0;
  if (core === 'drafts' || core === 'draft') return 1;
  if (core === 'sent items' || core === 'sent') return 2;
  if (core === 'deleted items' || core === 'trash' || core === 'deleted') return 3;
  if (core === 'junk email' || core === 'junk' || core === 'spam') return 4;
  if (core === 'archive' || core === 'archives') return 5;
  if (core === 'notes') return 6;
  if (core === 'outbox') return 7;

  return Number.POSITIVE_INFINITY; // user-defined folder
}

/**
 * Stable sort key for folders that share the same well-known rank:
 *   [rank, isTopLevel, originalIndex]
 * - Top-level folders come before their subfolders
 * - Otherwise preserve Graph's original order
 */
function folderSortKey(
  folder: GraphFolder,
  originalIndex: number
): [number, number, string] {
  const rank = wellKnownRank(folder);
  const isTopLevel = !folder.parentFolderId ? 0 : 1;
  return [rank, isTopLevel, originalIndex.toString().padStart(6, '0')];
}

/**
 * Transform multiple folders - build hierarchy and apply Outlook's
 * canonical folder ordering.
 *
 * Microsoft Graph returns folders in arbitrary order (often
 * ascending by `displayName`). The Outlook desktop / web client,
 * however, renders system folders in a fixed sequence:
 *   Inbox → Drafts → Sent Items → Deleted Items → Junk Email →
 *   Archive → Notes → Outbox → (user folders, then user subfolders
 *   in their declared order).
 *
 * Without an explicit sort the UI shows folders in odd
 * alphabetical order. This function sorts the flat list using
 * `wellKnownName` (with a display-name fallback) so the system
 * folders appear in Outlook's order, while still preserving
 * parent/child relationships.
 */
export function transformAndHierarchizeFolders(
  graphFolders: GraphFolder[]
): FrontendFolder[] {
  // First pass: transform all folders (icon, name, parent, etc.)
  const transformed = graphFolders.map((gf) =>
    transformGraphFolderToFrontend(gf, gf.unreadItemCount || 0)
  );

  // Pair with their original index so we keep a stable secondary
  // order for folders that share the same well-known rank.
  const indexed = transformed.map((folder, index) => ({
    folder,
    originalIndex: index,
    graph: graphFolders[index],
  }));

  indexed.sort((a, b) => {
    const [aRank, aTop, aIdx] = folderSortKey(a.graph, a.originalIndex);
    const [bRank, bTop, bIdx] = folderSortKey(b.graph, b.originalIndex);
    if (aRank !== bRank) return aRank - bRank;
    if (aTop !== bTop) return aTop - bTop;
    // Same rank + same top-level status: preserve Graph order.
    return aIdx.localeCompare(bIdx);
  });

  const sortedFolders = indexed.map((entry) => entry.folder);

  // Now restructure the tree so the layout matches Outlook:
  //
  //   ⭐ Favorites                    (synthetic group)
  //      📥 Inbox                     (shortcut — same real id)
  //      📤 Sent Items                (shortcut — same real id)
  //   📁 Folders                      (synthetic group)
  //      📥 Inbox                     (real folder)
  //      📝 Drafts                    (real folder)
  //      📤 Sent Items                (real folder)
  //      🗑️ Deleted Items             (real folder)
  //      ⚠️ Junk Email                (real folder)
  //      📦 Archive                   (real folder + subfolders)
  //      📒 Notes                     (real folder)
  //      📮 Outbox                    (real folder)
  //   📂 Work Projects                (user group)
  //      ...
  //
  // The Favorites shortcuts and the matching Folders entries share
  // the same real Graph folder id, so clicking either selects the
  // same real folder. We disambiguate the rendered rows by
  // appending `__fav_<rank>` to the shortcut's synthetic id and
  // setting a `realFolderId` field, then collapsing it in the
  // click handler.
  //
  // Out of the box Outlook's default favorites are Inbox, Drafts,
  // Sent Items, and Deleted Items. We replicate that here.
  const FAVORITE_RANKS: number[] = [0, 1, 2, 3]; // Inbox, Drafts, Sent Items, Deleted Items

  const favoritesGroup: FrontendFolder = {
    id: 'favorites',
    name: 'Favorites',
    hasChild: true,
    expanded: true,
    parentId: undefined,
  };

  const foldersGroup: FrontendFolder = {
    id: '__folders_root__',
    name: 'Folders',
    hasChild: true,
    expanded: true,
    parentId: undefined,
  };

  // Build the Favorites shortcuts. Each shortcut clones the real
  // folder so the icon/name/count come from Graph, but it points
  // at the synthetic Favorites group as its parent and uses a
  // derived id so the UI can treat it as a distinct node.
  const favoriteShortcuts: FrontendFolder[] = [];
  const seenFavRealIds = new Set<string>();
  for (let i = 0; i < sortedFolders.length; i++) {
    const folder = sortedFolders[i];
    // Only top-level system folders are eligible for Favorites.
    if (folder.parentId) continue;
    const graphEntry = graphFolders[indexed[i].originalIndex];
    if (!graphEntry) continue;
    const rank = wellKnownRank(graphEntry);
    if (!FAVORITE_RANKS.includes(rank)) continue;
    if (seenFavRealIds.has(folder.id)) continue;
    seenFavRealIds.add(folder.id);

    favoriteShortcuts.push({
      ...folder,
      id: getFolderShortcutId(folder.id),
      parentId: 'favorites',
      expanded: false,
    });
  }

  // Re-parent the well-known system folders so they live under the
  // synthetic "Folders" group instead of at the top level. Their
  // ids are preserved so messages still match by folderId.
  const SYSTEM_FOLDERS_GROUP_ID = '__folders_root__';
  const restFolders: FrontendFolder[] = sortedFolders.map((folder, i) => {
    // Subfolders (parentId set) keep their parent — we only re-home
    // top-level system folders.
    if (folder.parentId) return folder;

    // Top-level folder. If it's a well-known system folder, move
    // it under the Folders group.
    const graphEntry = graphFolders[indexed[i].originalIndex];
    const rank = graphEntry ? wellKnownRank(graphEntry) : Number.POSITIVE_INFINITY;
    if (rank === Number.POSITIVE_INFINITY) {
      // User-created top-level folder (e.g. "Work Projects"):
      // keep it as a root.
      return folder;
    }
    return {
      ...folder,
      parentId: SYSTEM_FOLDERS_GROUP_ID,
    };
  });

  // Final assembled tree, in the order Outlook displays it.
  return [
    favoritesGroup,
    ...favoriteShortcuts,
    foldersGroup,
    ...restFolders,
  ];
}

/**
 * Build the synthetic shortcut id used for a Favorites entry.
 * Centralised so the frontend's FolderNavigation can reverse the
 * mapping with `getRealFolderId`.
 */
export function getFolderShortcutId(realFolderId: string): string {
  return `__fav__${realFolderId}`;
}

/**
 * Reverse the shortcut mapping. If the id is a synthetic Favorites
 * shortcut (e.g. `__fav__AAMk...`), return the real Graph folder
 * id it represents. Otherwise return the id unchanged.
 */
export function getRealFolderId(id: string): string {
  if (typeof id === 'string' && id.startsWith('__fav__')) {
    return id.slice('__fav__'.length);
  }
  return id;
}

/**
 * Transform Microsoft Graph Message to Frontend MailData
 */
export interface FrontendMail {
  id: string;
  folderId: string;
  senderName: string;
  senderEmail: string;
  subject: string;
  snippet: string;
  date: string;
  isRead: boolean;
  isStarred: boolean;
  isFlagged: boolean;
  avatarColor: string;
  body?: string;
}

export interface GraphMessage {
  id: string;
  subject: string;
  bodyPreview: string;
  body?: { content: string };
  sender?: { emailAddress: { name: string; address: string } };
  from?: { emailAddress: { name: string; address: string } };
  receivedDateTime: string;
  isRead: boolean;
  categories?: string[];
  flag?: { flagStatus: string };
}

export function transformGraphMessageToFrontend(
  graphMessage: GraphMessage,
  folderId: string
): FrontendMail {
  // Extract sender info
  const senderInfo = graphMessage.sender?.emailAddress || graphMessage.from?.emailAddress;
  const senderName = senderInfo?.name || 'Unknown';
  const senderEmail = senderInfo?.address || '';

  // Check flags
  const isStarred = graphMessage.categories?.includes('flagged') || false;
  const isFlagged = graphMessage.flag?.flagStatus === 'flagged' || false;

  return {
    id: graphMessage.id,
    folderId,
    senderName,
    senderEmail,
    subject: graphMessage.subject || '(No Subject)',
    snippet: graphMessage.bodyPreview || '',
    // Pass the raw UTC ISO timestamp straight through. Microsoft
    // Graph returns `receivedDateTime` in UTC (e.g.
    // "2026-07-27T15:59:00Z"). Formatting it on the backend would
    // lock the display to the *server's* timezone, which may
    // differ from the user's — producing wrong times in the list.
    // The frontend (`formatMailDate` in
    // my-app/src/utils/transformers.ts) renders the date in the
    // browser's local timezone, which matches what Outlook on the
    // web shows for the signed-in user.
    date: graphMessage.receivedDateTime,
    isRead: graphMessage.isRead,
    isStarred,
    isFlagged,
    avatarColor: generateColorFromString(senderEmail),
    body: graphMessage.body?.content,
  };
}

/**
 * Transform batch of messages
 */
export function transformGraphMessagesToFrontend(
  graphMessages: GraphMessage[],
  folderId: string
): FrontendMail[] {
  return graphMessages.map((msg) => transformGraphMessageToFrontend(msg, folderId));
}

/**
 * Format response for initial data load
 */
export interface InitialDataResponse {
  folders: FrontendFolder[];
  mailsByFolder: { [folderId: string]: FrontendMail[] };
  totalMailCount: number;
}

export function formatInitialDataResponse(
  folders: FrontendFolder[],
  mailsByFolder: { [folderId: string]: FrontendMail[] }
): InitialDataResponse {
  const totalMailCount = Object.values(mailsByFolder).reduce(
    (sum, mails) => sum + mails.length,
    0
  );

  return {
    folders,
    mailsByFolder,
    totalMailCount,
  };
}
