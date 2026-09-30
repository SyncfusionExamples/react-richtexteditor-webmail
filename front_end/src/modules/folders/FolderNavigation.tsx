import React, { useMemo } from 'react';
import { useFolderStore } from '../../store/Store';
import type { FolderData } from '../../data/mockData';
import { Inbox, Send, FileText, Trash2, Archive, AlertCircle, ChevronRight, ChevronDown } from 'lucide-react';
import { FolderContextMenu } from '../../shared/components/FolderContextMenu';
import { getRealFolderId } from '../../utils/transformers';

const IconMap: Record<string, any> = {
  Inbox: Inbox,
  Send: Send,
  FileText: FileText,
  Trash2: Trash2,
  Archive: Archive,
  AlertCircle: AlertCircle,
};

// Memoized folder template component to prevent re-renders
const FolderTemplate = React.memo(({ data }: { data: FolderData }) => {
  const IconComponent = data.icon ? IconMap[data.icon] : null;
  return (
    <FolderContextMenu folder={data}>
      <div className="flex items-center justify-between w-full hover:bg-gray-100 rounded px-2 py-1 transition-colors gap-4">
        <div className="flex items-center gap-2.5 min-w-0">
          {IconComponent && <IconComponent size={16} className="text-gray-600 flex-shrink-0" />}
          <span className="text-sm text-gray-800 truncate">{data.name}</span>
        </div>
        {data.count !== undefined && data.count > 0 && (
          <span className="text-xs text-gray-500 font-medium flex-shrink-0">{data.count}</span>
        )}
      </div>
    </FolderContextMenu>
  );
});

FolderTemplate.displayName = 'FolderTemplate';

export const FolderNavigation: React.FC = () => {
  const { folders, selectedFolderId, setSelectedFolderId } = useFolderStore();

  // Track which folders are currently expanded. We seed the set
  // from each folder's `expanded` flag (Inbox starts expanded, as
  // it did in the original component). The set is keyed by the
  // folder's *real* id (after `getRealFolderId`) so Favorites
  // shortcuts and the real folders they point to share state.
  const [expandedIds, setExpandedIds] = React.useState<Set<string>>(() => {
    const seeded = new Set<string>();
    for (const f of folders) {
      // Only seed folders that are likely to be useful: ones
      // flagged as expanded, plus the synthetic groups so the
      // tree opens with Favorites + Folders already expanded.
      if (f.expanded || f.id === 'favorites' || f.id === '__folders_root__') {
        seeded.add(f.id);
      }
    }
    return seeded;
  });

  const toggleExpanded = (id: string) => {
    setExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const childrenByParentId = useMemo(() => {
    const map = new Map<string | undefined, FolderData[]>();

    for (const folder of folders) {
      const key = folder.parentId;
      const current = map.get(key) || [];
      current.push(folder);
      map.set(key, current);
    }

    return map;
  }, [folders]);

  const renderFolderTree = (folder: FolderData, depth: number) => {
    const children = childrenByParentId.get(folder.id) || [];
    // The Favorites shortcuts use a synthetic id like
    // `__fav__<realId>`. When the user clicks one we want the
    // selected folder to be the real one (so the mail list
    // re-fetches for the real Graph folder), but the row's own
    // highlight should still track the shortcut.
    const realId = getRealFolderId(folder.id);
    const isSelected = realId === selectedFolderId;
    const hasChildren = children.length > 0;
    // A folder is "open" if it's in the expanded set. We also
    // treat a folder as open by default if it has no record in
    // the set yet — this matches the seeded behavior for the
    // Favorites / Folders synthetic groups, so the very first
    // paint shows the tree expanded even before the useState
    // initializer runs (e.g. on HMR reloads).
    const isOpen = expandedIds.has(folder.id) || (
      !expandedIds.has(`__closed__${folder.id}`) &&
      (folder.id === 'favorites' || folder.id === '__folders_root__')
    );

    // Indentation strategy:
    //   - depth 0 (roots) get 12px left padding.
    //   - each child level adds 18px.
    //   - we also add a thin vertical guide line on the LEFT of
    //     every nested level so subfolders are visually obvious
    //     even if the user's stylesheet overrides padding.
    //   - a small chevron next to the icon shows the folder's
    //     expand/collapse state and is clickable to toggle it.
    const baseLeft = 12;
    const stepLeft = 18;
    const left = baseLeft + depth * stepLeft;

    return (
      <div
        key={folder.id}
        data-folder-depth={depth}
        data-folder-id={folder.id}
        data-folder-parent={folder.parentId ?? 'root'}
        data-folder-open={isOpen ? 'true' : 'false'}
        className="relative"
      >
        {/* Vertical guide line for nested levels. We use a thin
            border-l on a positioned wrapper so the line extends
            the full height of the row. */}
        {depth > 0 && (
          <div
            className="absolute top-0 bottom-0 border-l border-gray-200 pointer-events-none"
            style={{ left: `${baseLeft + (depth - 1) * stepLeft + 8}px` }}
            aria-hidden="true"
          />
        )}
        <FolderContextMenu folder={folder}>
          <div
            className={`relative w-full flex items-center justify-between gap-2 rounded py-1 pr-2 text-left transition-colors hover:bg-gray-100 ${isSelected ? 'bg-blue-50 text-blue-700' : ''}`}
            style={{ paddingLeft: `${left}px` }}
          >
            <div className="flex items-center gap-1.5 min-w-0 flex-1">
              {/* Expand/collapse chevron. Clicking it toggles the
                  folder's open state without selecting the folder.
                  Folders without children get a placeholder so all
                  rows line up vertically. */}
              {hasChildren ? (
                <button
                  type="button"
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleExpanded(folder.id);
                  }}
                  className="flex items-center justify-center w-3 h-3 flex-shrink-0 text-gray-400 hover:text-gray-700 focus:outline-none"
                  aria-label={isOpen ? `Collapse ${folder.name}` : `Expand ${folder.name}`}
                  aria-expanded={isOpen}
                  data-folder-toggle={folder.id}
                >
                  {isOpen ? (
                    <ChevronDown size={12} />
                  ) : (
                    <ChevronRight size={12} />
                  )}
                </button>
              ) : (
                <span className="w-3 h-3 flex-shrink-0" aria-hidden="true" />
              )}
              {/* The main click target — selecting a folder. */}
              <button
                type="button"
                onClick={() => setSelectedFolderId(realId)}
                className="flex items-center gap-1.5 min-w-0 flex-1 text-left focus:outline-none"
                data-folder-select={folder.id}
              >
                {folder.icon ? (() => {
                  const IconComponent = IconMap[folder.icon];
                  return IconComponent ? <IconComponent size={14} className="text-gray-600 flex-shrink-0" /> : null;
                })() : null}
                <span className="text-sm text-inherit truncate">{folder.name}</span>
              </button>
            </div>
            {folder.count !== undefined && folder.count > 0 && (
              <span className="text-xs text-gray-500 font-medium flex-shrink-0">{folder.count}</span>
            )}
          </div>
        </FolderContextMenu>

        {hasChildren && isOpen && (
          <div className="mt-0.5">
            {children.map((child) => renderFolderTree(child, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  const rootFolders = useMemo(() => {
    const knownFolderIds = new Set(folders.map(folder => folder.id));
    // A folder is rendered as a root when:
    //   - it has no `parentId` (a true top-level folder), OR
    //   - its `parentId` points to a folder we don't have in the
    //     list (an orphan whose real parent is missing — render it
    //     at the top level rather than hiding it).
    //
    // We intentionally do NOT treat `parentId === 'favorites'` as a
    // root any more. The synthetic Favorites shortcuts have
    // `parentId: 'favorites'` AND should be children of the
    // Favorites group. Previously the filter also let them pass as
    // roots, which made them render twice in the tree (once at the
    // top, once under Favorites) and, depending on the synthetic
    // group order, could leave the Favorites group looking empty
    // when the shortcuts were filtered out as "already a child".
    return folders.filter(
      (folder) =>
        !folder.parentId || !knownFolderIds.has(folder.parentId)
    );
  }, [folders]);

  return (
    <div className="h-full overflow-y-auto overflow-x-hidden py-4">
      <div className="space-y-0.5">
        {rootFolders.map((folder) => renderFolderTree(folder, 0))}
      </div>
    </div>
  );
};
