/**
 * Transform Microsoft Graph API responses to match frontend expectations
 */

// Color palette for avatar generation
const COLORS = [
  '#0078d4', '#107c10', '#d13438', '#4c4a48', '#ffc646', '#7fba00', '#7c4a7e', '#f7630c'
];

/**
 * Format a mail date for display in the UI.
 *
 * Microsoft Graph returns dates as ISO 8601 UTC strings (e.g.
 * "2026-07-27T09:00:34Z"). We render them in the user's local
 * timezone (the browser's), matching what Outlook on the web
 * shows for the signed-in user.
 *
 * Two modes:
 *   - "list" (default): Outlook's mail-list format.
 *       * Today        → "3:59 PM"
 *       * Same year    → "Mon 7/27"
 *       * Older year   → "Mon 7/27/2024"
 *   - "full": the mail-reader header format, used when a mail is
 *     opened. Always shows the full date+time, matching Outlook's
 *     reader pane:
 *       → "Mon 7/27/2026 3:59 PM"
 *
 * Falls back to the original string if the value can't be parsed
 * (e.g. a non-ISO value already produced by local-only drafts).
 */
export function formatMailDate(
  value: string | Date | undefined | null,
  mode: 'list' | 'full' = 'full'
): string {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return typeof value === 'string' ? value : '';
  }

  const weekday = date.toLocaleDateString('en-US', { weekday: 'short' });
  const month = date.getMonth() + 1; // 1-12, no leading zero (matches M/D/YYYY)
  const day = date.getDate();        // 1-31, no leading zero
  const year = date.getFullYear();
  let hour = date.getHours();        // 0-23
  const minute = date.getMinutes().toString().padStart(2, '0');
  const ampm = hour >= 12 ? 'PM' : 'AM';
  hour = hour % 12;
  if (hour === 0) hour = 12;

  if (mode === 'list') {
    const now = new Date();
    const isSameDay =
      date.getFullYear() === now.getFullYear() &&
      date.getMonth() === now.getMonth() &&
      date.getDate() === now.getDate();
    const isSameYear = date.getFullYear() === now.getFullYear();

    if (isSameDay) {
      return `${hour}:${minute} ${ampm}`;
    }
    if (isSameYear) {
      return `${weekday} ${month}/${day}`;
    }
    return `${weekday} ${month}/${day}/${year}`;
  }

  return `${weekday} ${month}/${day}/${year} ${hour}:${minute} ${ampm}`;
}

/**
 * Map a mail's date to an Outlook-style date group label used as a
 * section header in the mail list.
 *
 * Groups (in the order they appear in the list, newest first):
 *   - "Today"
 *   - "Yesterday"
 *   - Day-of-week name (e.g. "Monday", "Sunday") — for mails within
 *     the last 7 days that aren't Today / Yesterday
 *   - "Last Week"
 *   - "Two Weeks Ago"
 *   - "Three Weeks Ago"
 *   - "Earlier This Month"
 *   - "Last Month"
 *   - "Older"
 *
 * Falls back to "Older" if the date can't be parsed.
 */
export function getMailDateGroup(value: string | Date | undefined | null): string {
  if (!value) return 'Older';
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return typeof value === 'string' ? value : 'Older';
  }

  const now = new Date();
  // Use local-midnight-based diffs so "today" matches the user's
  // calendar, not a 24h rolling window.
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  const dayDiff = Math.round(
    (startOfToday.getTime() - startOfDate.getTime()) / (24 * 60 * 60 * 1000)
  );

  if (dayDiff <= 0) return 'Today';
  if (dayDiff === 1) return 'Yesterday';
  if (dayDiff >= 2 && dayDiff <= 6) {
    // Within the last week but not today/yesterday — named day.
    return date.toLocaleDateString('en-US', { weekday: 'long' });
  }
  if (dayDiff >= 7 && dayDiff <= 13) return 'Last Week';
  if (dayDiff >= 14 && dayDiff <= 20) return 'Two Weeks Ago';
  if (dayDiff >= 21 && dayDiff <= 27) return 'Three Weeks Ago';

  // Once we're past ~3 weeks, group by month relative to today.
  const isSameMonth =
    date.getFullYear() === now.getFullYear() &&
    date.getMonth() === now.getMonth();
  if (isSameMonth) return 'Earlier This Month';

  // One calendar month back from today's month (across year
  // boundary if needed).
  const lastMonthDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const isLastMonth =
    date.getFullYear() === lastMonthDate.getFullYear() &&
    date.getMonth() === lastMonthDate.getMonth();
  if (isLastMonth) return 'Last Month';

  return 'Older';
}

/**
 * Ordered list of Outlook-style date-group labels, used by the
 * mail list so groups render in the same order Outlook uses
 * (Today -> Yesterday -> named day -> Last Week -> ... -> Older)
 * regardless of the order they appear in the data.
 */
export const MAIL_DATE_GROUP_ORDER: string[] = [
  'Today',
  'Yesterday',
  // Named days of the week — we drop these in dynamically based on
  // the actual days encountered, so they stay in the right slot
  // (after Yesterday, before Last Week). The list below is just
  // the canonical weekday ordering.
  'Sunday',
  'Monday',
  'Tuesday',
  'Wednesday',
  'Thursday',
  'Friday',
  'Saturday',
  'Last Week',
  'Two Weeks Ago',
  'Three Weeks Ago',
  'Earlier This Month',
  'Last Month',
  'Older',
];

/**
 * Sort key for a date-group label, so groups render in Outlook's
 * order. Groups not in the canonical list (e.g. weekday names)
 * are placed between "Yesterday" and "Last Week", ordered by
 * RECENCY rather than by the weekday's calendar position.
 *
 * Background: a weekday name like "Monday" or "Wednesday" on its
 * own is ambiguous — Monday is the most recent day only when
 * today is Tuesday. Previously this function used the weekday's
 * calendar index (Sunday=0..Saturday=6), which meant that when
 * today was Friday the list rendered as Today → Yesterday →
 * Monday → Wednesday (because Monday's index 1 < Wednesday's
 * index 3). The correct chronological order is Today → Yesterday
 * → Wednesday → Tuesday → Monday → Sunday → Saturday.
 *
 * To produce that ordering we anchor weekday ranks to "yesterday":
 * the day immediately before yesterday is closest to 1.5 (just
 * after Yesterday's rank of 1), and each earlier weekday is
 * offset further by 1/7. This guarantees weekday groups always
 * appear in true recency order, wrapping around the week boundary
 * correctly.
 */
export function dateGroupRank(label: string): number {
  // IMPORTANT: weekday names ("Sunday", "Monday", ...) are present
  // in MAIL_DATE_GROUP_ORDER as placeholders, but their ordering
  // there is the calendar order (Sun..Sat) which does NOT match
  // the chronological order we want. The correct order for the
  // mail list is "most-recent weekday first" relative to today —
  // e.g. when today is Friday the list should be Today, Yesterday,
  // Wednesday, Tuesday, Monday, Sunday, Saturday — NOT
  // Today, Yesterday, Sunday, Monday, Tuesday, Wednesday.
  //
  // We therefore special-case weekday names BEFORE the canonical
  // lookup, so they always go through the recency-based branch
  // and ignore their placeholder index in MAIL_DATE_GROUP_ORDER.
  const weekdayOrder = [
    'Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday',
  ];
  const wd = weekdayOrder.indexOf(label);
  if (wd >= 0) {
    // Anchor: yesterday is at rank 1 in MAIL_DATE_GROUP_ORDER.
    // We compute yesterday's weekday index (today - 1, with wrap)
    // and then walk backward day-by-day, incrementing the rank
    // by 1/7 per step. The first step after yesterday is the
    // most recent "named day" group, the next step is the day
    // before that, and so on, wrapping Sunday→Saturday as needed.
    const todayWd = new Date().getDay(); // 0=Sun..6=Sat
    const yesterdayWd = (todayWd - 1 + 7) % 7;
    // Number of days from "yesterday" to the target weekday,
    // walking backward in time (and wrapping around the week).
    // For example, if yesterday is Thursday (wd=4):
    //   Wednesday (3) → 1 step back
    //   Tuesday  (2)  → 2 steps back
    //   Monday   (1)  → 3 steps back
    //   Sunday   (0)  → 4 steps back
    //   Saturday (6)  → 5 steps back
    const stepsBack = (yesterdayWd - wd + 7) % 7;
    // stepsBack is 0 for "yesterday" itself, but yesterday is
    // already handled by MAIL_DATE_GROUP_ORDER above. For any
    // other weekday stepsBack is 1..6, and the rank is
    //   1 + stepsBack/7
    // which always lands strictly between Yesterday (1) and the
    // first canonical fixed label after it (Last Week at 9).
    return 1 + stepsBack / 7;
  }
  // Non-weekday label — fall back to its position in the canonical
  // ordering (Today, Yesterday, Last Week, ..., Older).
  const idx = MAIL_DATE_GROUP_ORDER.indexOf(label);
  if (idx >= 0) return idx;
  return Number.POSITIVE_INFINITY;
}

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
   * Graph-assigned well-known folder identifier. NOTE: this field
   * is NOT requested in the Graph `select` (some tenants reject it
   * with: "Could not find a property named 'wellKnownName' on type
   * 'microsoft.graph.mailFolder'"), so it will be `undefined` in
   * practice. The canonical Outlook ordering is resolved entirely
   * from `displayName` + `parentFolderId` in
   * `transformAndHierarchizeFolders`.
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
 * `wellKnownName` first, then falls back to a display-name match
 * so the UI still gets Outlook-style ordering when Graph omits
 * `wellKnownName` (older tenants, custom renames, etc.).
 */
function wellKnownRank(folder: GraphFolder): number {
  // Prefer `wellKnownName` if Graph ever provides it (currently we
  // don't request it because some tenants reject the field in a
  // flat `select`, but if it shows up the lookup is unambiguous).
  const wk = (folder.wellKnownName || '').toLowerCase();
  if (wk && WELL_KNOWN_ORDER[wk] !== undefined) {
    return WELL_KNOWN_ORDER[wk];
  }

  // Fall back to a localized display-name match. We require the
  // distinguishing token to be the WHOLE name (or the whole name plus
  // a parenthetical like "Sent Items (Archive)") to avoid colliding
  // with user folders that happen to contain a generic word. The
  // top-level system folders Graph returns have predictable English
  // labels: "Inbox", "Drafts", "Sent Items", "Deleted Items",
  // "Junk Email", "Archive", "Notes", "Outbox".
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
function folderSortKey(folder: GraphFolder, originalIndex: number): [number, number, string] {
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
 * alphabetical order. This function:
 *   1. Sorts the flat list using `wellKnownName` (with a
 *      display-name fallback) so the system folders appear in
 *      Outlook's order.
 *   2. Preserves parent/child relationships — subfolders render
 *      under their real parent in the frontend tree.
 *   3. Removes the previous synthetic "Favorites" root, which
 *      was empty and was hiding the actual system folders.
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

  // Wrap the tree in Outlook-style synthetic groups so the layout
  // matches what the user sees in Outlook:
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
  // Favorites shortcuts share the real Graph folder id (cloned into
  // a derived id so the UI can treat them as distinct rows). The
  // click handler in FolderNavigation uses `getRealFolderId` to map
  // the shortcut row back to the real folder id when selected.
  const FAVORITE_RANKS: number[] = [0, 1, 2, 3]; // Inbox, Drafts, Sent Items, Deleted Items
  const SYSTEM_FOLDERS_GROUP_ID = '__folders_root__';

  const favoritesGroup: FrontendFolder = {
    id: 'favorites',
    name: 'Favorites',
    hasChild: true,
    expanded: true,
    parentId: undefined,
  };

  const foldersGroup: FrontendFolder = {
    id: SYSTEM_FOLDERS_GROUP_ID,
    name: 'Folders',
    hasChild: true,
    expanded: true,
    parentId: undefined,
  };

  // Build Favorites shortcuts by cloning each eligible system
  // folder. We re-rank using the same `wellKnownRank` the sorter
  // used to keep the shortcut order matching Outlook's favorites
  // row order.
  const favoriteShortcuts: FrontendFolder[] = [];
  const seenFavRealIds = new Set<string>();
  for (let i = 0; i < sortedFolders.length; i++) {
    const folder = sortedFolders[i];
    if (folder.parentId) continue; // subfolders never shortcut
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

  // Re-parent well-known system folders under the synthetic Folders
  // group. User-created top-level folders stay at the top level
  // (matching the mockData layout where "Work Projects", "Personal",
  // etc. are roots). Subfolders keep their real parent.
  const restFolders: FrontendFolder[] = sortedFolders.map((folder, i) => {
    if (folder.parentId) return folder;
    const graphEntry = graphFolders[indexed[i].originalIndex];
    const rank = graphEntry ? wellKnownRank(graphEntry) : Number.POSITIVE_INFINITY;
    if (rank === Number.POSITIVE_INFINITY) return folder;
    return {
      ...folder,
      parentId: SYSTEM_FOLDERS_GROUP_ID,
    };
  });

  return [
    favoritesGroup,
    ...favoriteShortcuts,
    foldersGroup,
    ...restFolders,
  ];
}

/**
 * Build the synthetic shortcut id used for a Favorites entry.
 * Centralised so the click handler in FolderNavigation can reverse
 * the mapping with `getRealFolderId`.
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
 * Rank a folder's display name against the Outlook well-known
 * order. Used by `wrapWithOutlookGroups` (which doesn't have the
 * original Graph entries to call `wellKnownRank` on) to decide
 * which folders go into the Favorites group and which into the
 * Folders group.
 *
 * Returns the same ranks as `wellKnownRank`:
 *   0=Inbox, 1=Drafts, 2=Sent, 3=Deleted, 4=Junk,
 *   5=Archive, 6=Notes, 7=Outbox,
 *   Number.POSITIVE_INFINITY = user folder.
 */
function wellKnownNameRank(name: string): number {
  const core = (name || '').toLowerCase().trim().replace(/\s*\([^)]*\)\s*$/, '').trim();
  if (core === 'inbox') return 0;
  if (core === 'drafts' || core === 'draft') return 1;
  if (core === 'sent items' || core === 'sent') return 2;
  if (core === 'deleted items' || core === 'trash' || core === 'deleted') return 3;
  if (core === 'junk email' || core === 'junk' || core === 'spam') return 4;
  if (core === 'archive' || core === 'archives') return 5;
  if (core === 'notes') return 6;
  if (core === 'outbox') return 7;
  return Number.POSITIVE_INFINITY;
}

/**
 * Wrap an already-transformed flat folder list with the synthetic
 * Outlook Favorites + Folders groups, in canonical order.
 *
 * This is a defensive helper used by the frontend's `loadRealData`
 * so the UI shows the correct tree even if the backend response
 * came from an older build that didn't emit the synthetic groups.
 * The function is idempotent: if the input already contains a
 * `favorites` or `__folders_root__` group, the caller's data is
 * returned unchanged (the caller should check first).
 */
export function wrapWithOutlookGroups(
  folders: FrontendFolder[]
): FrontendFolder[] {
  // First, strip out any pre-existing synthetic nodes from the
  // input. We rebuild the Outlook tree from the real folders
  // only.
  //
  // We intentionally no longer emit a synthetic `favorites`
  // Favorites group: Microsoft Graph doesn't expose the user's
  // Outlook Favorites list (those are stored in the user's
  // mailbox profile, not in the mailFolders hierarchy), so any
  // "Favorites" group we synthesized was empty of the user's
  // actual pinned folders and confused the user. Real Graph
  // folders now render directly under the synthetic Folders
  // group, matching what Graph actually returns.
  const cleaned: FrontendFolder[] = [];
  for (const folder of folders) {
    // Drop the synthetic group nodes.
    if (folder.id === 'favorites' || folder.id === '__folders_root__') continue;
    // Drop any pre-existing Favorites shortcuts (now stale).
    if (typeof folder.id === 'string' && folder.id.startsWith('__fav__')) continue;
    // Drop any folder that was previously pinned to Favorites via
    // `addToFavorites` (parentId 'favorites' or the legacy '1').
    // The pinned row is a duplicate of a real folder; the real
    // folder still renders under its true parent below.
    if (folder.parentId === 'favorites' || folder.parentId === '1') continue;
    cleaned.push(folder);
  }

  // Sort the cleaned list by well-known rank so the wrapped output
  // is in Outlook's order even if the caller skipped the sort.
  const sorted = [...cleaned].sort((a, b) => {
    const aRank = wellKnownNameRank(a.name);
    const bRank = wellKnownNameRank(b.name);
    if (aRank !== bRank) return aRank - bRank;
    // Within the same rank, top-level first.
    const aTop = !a.parentId ? 0 : 1;
    const bTop = !b.parentId ? 0 : 1;
    if (aTop !== bTop) return aTop - bTop;
    return 0;
  });

  const SYSTEM_FOLDERS_GROUP_ID = '__folders_root__';

  const foldersGroup: FrontendFolder = {
    id: SYSTEM_FOLDERS_GROUP_ID,
    name: 'Folders',
    hasChild: true,
    expanded: true,
    parentId: undefined,
  };

  // Re-parent the well-known system folders under the synthetic
  // Folders group. User-created top-level folders stay at the root
  // so "Work Projects", "Personal", etc. still appear as siblings
  // of the Folders group. Subfolders keep their real parent.
  //
  // We also do a robust integrity pass: every subfolder's
  // `parentId` must reference a folder that's actually in the
  // list. If a parent is missing (e.g. the backend fetched the
  // child but failed to fetch the parent because of a permission
  // error, or the data is from an older build, or the parent's
  // id was renamed in transit), we attach the orphan to the
  // synthetic Folders group so it appears nested under the
  // system folders, never as a flat root. This guarantees
  // subfolders always render with the expected indentation.
  const folderById = new Map(sorted.map((f) => [f.id, f]));
  const restFolders: FrontendFolder[] = sorted.map((folder) => {
    if (folder.parentId) {
      // Subfolder: validate that the immediate parent is in the
      // list. We don't walk further up because each link in the
      // chain must be present in the data for the tree to render
      // correctly — if any link is missing, the safest fallback
      // is to re-attach the subfolder to the synthetic Folders
      // group so it's at least indented.
      const hasParent = folderById.has(folder.parentId);
      if (hasParent) {
        return folder;
      }
      // Couldn't resolve. Attach to the synthetic Folders group
      // so the subfolder is at least nested and indented.
      console.warn(
        `wrapWithOutlookGroups: folder "${folder.name}" (${folder.id}) has parentId "${folder.parentId}" but no matching parent in the list. Attaching to Folders group.`
      );
      return {
        ...folder,
        parentId: SYSTEM_FOLDERS_GROUP_ID,
      };
    }
    // Top-level folder. Re-home well-known system folders under
    // the synthetic Folders group; leave user-created roots alone.
    const rank = wellKnownNameRank(folder.name);
    if (rank === Number.POSITIVE_INFINITY) return folder;
    return {
      ...folder,
      parentId: SYSTEM_FOLDERS_GROUP_ID,
    };
  });

  return [foldersGroup, ...restFolders];
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

  // Format date
  const date = new Date(graphMessage.receivedDateTime);
  const today = new Date();
  let dateStr = '';

  if (date.toDateString() === today.toDateString()) {
    dateStr = date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  } else if (date.getFullYear() === today.getFullYear()) {
    dateStr = date.toLocaleDateString([], { month: 'short', day: 'numeric' });
  } else {
    dateStr = date.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' });
  }

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
    date: dateStr,
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
