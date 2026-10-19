import React, { useRef } from 'react';
import { ContextMenuComponent } from '@syncfusion/ej2-react-navigations';
import type { MenuItemModel } from '@syncfusion/ej2-navigations';
import { Filter } from 'lucide-react';
import { useUIStore } from '../../store/Store';

export const FilterMenu: React.FC = () => {
  const { filters, setFilters } = useUIStore();
  const contextMenuRef = useRef<ContextMenuComponent>(null);

  const menuItems: MenuItemModel[] = [
    { 
      text: 'Read Status',
      items: [
        {
          text: 'All',
          id: 'read-all',
          iconCss: filters.isRead === null ? 'e-icons e-check' : ''
        },
        {
          text: 'Read',
          id: 'read-true',
          iconCss: filters.isRead === true ? 'e-icons e-check' : ''
        },
        {
          text: 'Unread',
          id: 'read-false',
          iconCss: filters.isRead === false ? 'e-icons e-check' : ''
        }
      ]
    },
    { 
      text: 'Flag Status',
      items: [
        {
          text: 'All',
          id: 'flag-all',
          iconCss: filters.isFlagged === null ? 'e-icons e-check' : ''
        },
        {
          text: 'Flagged',
          id: 'flag-true',
          iconCss: filters.isFlagged === true ? 'e-icons e-check' : ''
        },
        {
          text: 'Unflagged',
          id: 'flag-false',
          iconCss: filters.isFlagged === false ? 'e-icons e-check' : ''
        }
      ]
    },
    { separator: true },
    {
      text: 'Clear All Filters',
      id: 'clear-all'
    }
  ];

  const handleSelect = (args: any) => {
    const itemId = args.item.id;

    if (itemId === 'clear-all') {
      setFilters({});
      return;
    }

    if (itemId.startsWith('read-')) {
      const value = itemId === 'read-all' ? null : itemId === 'read-true';
      setFilters({ ...filters, isRead: value });
    } else if (itemId.startsWith('flag-')) {
      const value = itemId === 'flag-all' ? null : itemId === 'flag-true';
      setFilters({ ...filters, isFlagged: value });
    }
  };

  const handleFilterClick = (e: React.MouseEvent<HTMLButtonElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();

    // Position below the button (like dropdown)
    const top = rect.bottom + window.scrollY;
    const left = rect.left + window.scrollX;

    contextMenuRef.current?.open(top, left);

  };

  const hasActiveFilters = filters.isRead !== null || filters.isFlagged !== null;

  return (
    <>
      <button
        onClick={handleFilterClick}
        className={`p-2 rounded cursor-pointer border-none flex items-center gap-2 transition-colors ${
          hasActiveFilters
            ? 'bg-blue-100 text-blue-600 hover:bg-blue-200'
            : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
        }`}
        title="Filter mail"
      >
        <Filter size={20} />
        {hasActiveFilters && <span className="text-xs font-semibold">Filter</span>}
      </button>

      <ContextMenuComponent
        ref={contextMenuRef}
        items={menuItems}
        select={handleSelect}
        cssClass="custom-filter-menu"
      />
    </>
  );
};
