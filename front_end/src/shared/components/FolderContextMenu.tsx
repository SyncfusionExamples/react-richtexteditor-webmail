import React, { useRef, useState } from 'react';
import { ContextMenuComponent } from '@syncfusion/ej2-react-navigations';
import type { MenuItemModel as ContextMenuItemModel } from '@syncfusion/ej2-react-navigations';
import { useFolderStore } from '../../store/Store';
import type { FolderData } from '../../data/mockData';

interface FolderContextMenuProps {
  folder: FolderData;
  children: React.ReactNode;
}

export const FolderContextMenu: React.FC<FolderContextMenuProps> = ({ folder, children }) => {
  const { renameFolder, deleteFolder, createSubfolder, markFolderAsRead } = useFolderStore();
  const contextMenuRef = useRef<ContextMenuComponent>(null);
  const [showRenameDialog, setShowRenameDialog] = useState(false);
  const [showSubfolderDialog, setShowSubfolderDialog] = useState(false);
  const [newFolderName, setNewFolderName] = useState(folder.name);
  const [subfolderName, setSubfolderName] = useState('');

  // Root folders (no parentId) - check if this is a root level folder
  const isRootFolder = !folder.parentId;

  // We no longer render a synthetic Favorites group because
  // Microsoft Graph doesn't expose Outlook's per-user Favorites
  // list. The context menu therefore doesn't offer "Add to
  // Favorites" / "Remove from Favorites" any more — there's no
  // Favorites group to pin to.

  const menuItems: ContextMenuItemModel[] = [];

  // Basic actions available for all folders
  menuItems.push({
    text: 'Mark All as Read',
    id: 'mark-read',
    iconCss: 'e-icons e-checkbox-checked'
  });

  // Separator
  menuItems.push({ separator: true } as any);

  // New subfolder action (available for all real folders)
  menuItems.push({
    text: 'New Subfolder',
    id: 'new-subfolder',
    iconCss: 'e-icons e-folder-add'
  });

  // Actions only for non-root custom folders. Suppress rename /
  // delete on the synthetic Folders group itself so the user can't
  // accidentally remove the group node.
  if (!isRootFolder && folder.id !== '__folders_root__') {
    menuItems.push({ separator: true } as any);
    menuItems.push({
      text: 'Rename',
      id: 'rename-folder',
      iconCss: 'e-icons e-edit'
    });
    menuItems.push({
      text: 'Delete',
      id: 'delete-folder',
      iconCss: 'e-icons e-delete'
    });
  }

  const handleSelect = (args: any) => {
    const itemId = args.item.id;

    if (itemId === 'rename-folder') {
      setNewFolderName(folder.name);
      setShowRenameDialog(true);
    } else if (itemId === 'delete-folder') {
      if (window.confirm(`Are you sure you want to delete "${folder.name}"?`)) {
        deleteFolder(folder.id);
      }
    } else if (itemId === 'new-subfolder') {
      setSubfolderName('');
      setShowSubfolderDialog(true);
    } else if (itemId === 'mark-read') {
      markFolderAsRead(folder.id);
    }
  };

  const handleRenameSubmit = () => {
    if (newFolderName.trim() && newFolderName !== folder.name) {
      renameFolder(folder.id, newFolderName.trim());
    }
    setShowRenameDialog(false);
  };

  const handleRenameCancel = () => {
    setShowRenameDialog(false);
    setNewFolderName(folder.name);
  };

  const handleSubfolderSubmit = () => {
    if (subfolderName.trim()) {
      createSubfolder(folder.id, subfolderName.trim());
      setShowSubfolderDialog(false);
      setSubfolderName('');
    }
  };

  const handleSubfolderCancel = () => {
    setShowSubfolderDialog(false);
    setSubfolderName('');
  };

  const handleContextMenu = (e: React.MouseEvent<HTMLDivElement>) => {
    e.preventDefault();
    e.stopPropagation();
    contextMenuRef.current?.open(e.pageY, e.pageX);
  };

  return (
    <>
      <div onContextMenu={handleContextMenu}>
        {children}
        <ContextMenuComponent
          ref={contextMenuRef}
          items={menuItems}
          select={handleSelect}
          cssClass="custom-folder-context-menu"
        />
      </div>

      {/* Rename Dialog */}
      {showRenameDialog && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-lg p-6 w-96">
            <h2 className="text-lg font-semibold mb-4 text-gray-800">Rename Folder</h2>
            <input
              type="text"
              value={newFolderName}
              onChange={(e) => setNewFolderName(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded mb-4 focus:outline-none focus:border-blue-500"
              placeholder="Enter new folder name"
              autoFocus
              onKeyDown={(e) => {
                e.stopPropagation();
                if (e.key === 'Enter') {
                  handleRenameSubmit();
                }
              }}
            />
            <div className="flex gap-2 justify-end">
              <button
                onClick={handleRenameCancel}
                className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleRenameSubmit}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded cursor-pointer border-none transition-colors"
              >
                Rename
              </button>
            </div>
          </div>
        </div>
      )}

      {/* New Subfolder Dialog */}
      {showSubfolderDialog && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className="bg-white rounded-lg shadow-lg p-6 w-96">
            <h2 className="text-lg font-semibold mb-4 text-gray-800">Create New Subfolder</h2>
            <input
              type="text"
              value={subfolderName}
              onChange={(e) => setSubfolderName(e.target.value)}
              className="w-full px-3 py-2 border border-gray-300 rounded mb-4 focus:outline-none focus:border-blue-500"
              placeholder="Enter subfolder name"
              autoFocus
              onKeyDown={(e) => {
                e.stopPropagation();
                if (e.key === 'Enter') {
                  handleSubfolderSubmit();
                }
              }}
            />
            <div className="flex gap-2 justify-end">
              <button
                onClick={handleSubfolderCancel}
                className="px-4 py-2 text-gray-600 hover:bg-gray-100 rounded cursor-pointer border-none bg-transparent transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={handleSubfolderSubmit}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded cursor-pointer border-none transition-colors"
              >
                Create
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
};
