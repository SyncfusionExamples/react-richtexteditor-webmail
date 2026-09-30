import React, { useMemo, useCallback, useState } from 'react';
import { useFolderStore, useMailStore, useUIStore } from '../../store/Store';
import type { MailData } from '../../data/mockData';
import { Star, Flag, Eye, EyeOff, ChevronDown, ChevronRight } from 'lucide-react';
import { FilterMenu } from '../../shared/components/FilterMenu';
import {
  formatMailDate,
  getMailDateGroup,
  dateGroupRank,
} from '../../utils/transformers';

// Memoized mail list item component to prevent re-renders
const MailListItem = React.memo(({ 
  data, 
  isSelected, 
  onSelect, 
  onFlagClick, 
  onReadClick 
}: { 
  data: MailData;
  isSelected: boolean;
  onSelect: (id: string) => void;
  onFlagClick: (id: string, e: React.MouseEvent) => void;
  onReadClick: (id: string, e: React.MouseEvent) => void;
}) => (
  <div 
    className={`p-3 px-4 flex flex-col gap-1 cursor-pointer border-b border-gray-200 transition-colors ${isSelected ? 'bg-gray-100' : 'bg-transparent hover:bg-gray-50'}`}
    onClick={() => onSelect(data.id)}
  >
    <div className="flex justify-between items-center">
      <div className="flex items-center gap-2">
        {!data.isRead && <div className="w-2 h-2 rounded-full bg-blue-600" />}
        <span className={`text-sm text-gray-800 ${data.isRead ? 'font-normal' : 'font-semibold'}`}>{data.senderName}</span>
      </div>
      <span className="text-xs text-gray-600">{formatMailDate(data.date, 'list')}</span>
    </div>
    <div className="flex justify-between items-center">
      <span className={`text-[13px] text-gray-800 whitespace-nowrap overflow-hidden text-ellipsis max-w-[70%] ${data.isRead ? 'font-normal' : 'font-semibold'}`}>{data.subject}</span>
      <div className="flex items-center gap-1 ml-2">
        {data.isStarred && <Star size={14} fill="#fadb14" color="#fadb14" className="flex-shrink-0" />}
        <button
          onClick={(e) => onReadClick(data.id, e)}
          className={`p-1 rounded transition-colors cursor-pointer border-none bg-transparent flex items-center justify-center flex-shrink-0 ${data.isRead ? 'hover:bg-gray-200' : 'hover:bg-gray-200'}`}
          title={data.isRead ? 'Mark as unread' : 'Mark as read'}
        >
          {data.isRead ? (
            <Eye size={14} fill="#d1d5db" color="#d1d5db" className="cursor-pointer" />
          ) : (
            <EyeOff size={14} fill="#d1d5db" color="#d1d5db" className="cursor-pointer" />
          )}
        </button>
        <button
          onClick={(e) => onFlagClick(data.id, e)}
          className="p-1 hover:bg-gray-200 rounded transition-colors cursor-pointer border-none bg-transparent flex items-center justify-center flex-shrink-0"
          title={data.isFlagged ? 'Unflag' : 'Flag'}
        >
          <Flag 
            size={14} 
            fill={data.isFlagged ? "#ef4444" : "none"} 
            color={data.isFlagged ? "#ef4444" : "#ccc"}
            className="cursor-pointer"
          />
        </button>
      </div>
    </div>
    <div className="text-[13px] text-gray-600 whitespace-nowrap overflow-hidden text-ellipsis">
      {data.snippet}
    </div>
  </div>
));

MailListItem.displayName = 'MailListItem';

export const MailList: React.FC = () => {
  const { selectedFolderName } = useFolderStore();
  const { selectedFolderMails, selectedMailId, setSelectedMailId, toggleFlagStatus, toggleReadStatus, isLoading } = useMailStore();
  const { searchQuery, filters } = useUIStore();

  const filteredMails = useMemo(() => {
    let result = selectedFolderMails;

    // Apply search filter
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      result = result.filter(mail =>
        mail.senderName.toLowerCase().includes(query) ||
        mail.subject.toLowerCase().includes(query) ||
        mail.snippet.toLowerCase().includes(query)
      );
    }

    // Apply read status filter
    if (filters.isRead !== null && filters.isRead !== undefined) {
      result = result.filter(mail => mail.isRead === filters.isRead);
    }

    // Apply flag status filter
    if (filters.isFlagged !== null && filters.isFlagged !== undefined) {
      result = result.filter(mail => mail.isFlagged === filters.isFlagged);
    }

    return result;
  }, [selectedFolderMails, searchQuery, filters]);

  const handleFlagClick = useCallback((id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    toggleFlagStatus(id);
  }, [toggleFlagStatus]);

  const handleReadClick = useCallback((id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    toggleReadStatus(id);
  }, [toggleReadStatus]);

  const handleMailSelect = useCallback((id: string) => {
    setSelectedMailId(id);
  }, [setSelectedMailId]);

  // Track which date-group headers are currently collapsed. A
  // group is "collapsed" when its label is present in this set;
  // clicking the header toggles membership. State is keyed by
  // group label (e.g. "Today", "Yesterday", "Monday", "Older")
  // so it persists across folder switches and re-renders — if the
  // user collapses "Older" in Sent Items and then opens Inbox,
  // "Older" stays collapsed when they return to Sent Items.
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(
    () => new Set()
  );

  const toggleGroup = useCallback((groupLabel: string) => {
    setCollapsedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(groupLabel)) next.delete(groupLabel);
      else next.add(groupLabel);
      return next;
    });
  }, []);

  // Group mails by Outlook-style date groups. We compute this from
  // `filteredMails` so the search/flag/read filters flow through.
  // Mails are sorted in descending date order within each group,
  // and groups are sorted in Outlook's canonical order (Today ->
  // Yesterday -> named days -> Last Week -> ... -> Older).
  const groupedMails = useMemo(() => {
    const groups = new Map<string, MailData[]>();
    for (const mail of filteredMails) {
      const group = getMailDateGroup(mail.date);
      const list = groups.get(group) || [];
      list.push(mail);
      groups.set(group, list);
    }
    // Sort mails within each group newest-first.
    for (const list of groups.values()) {
      list.sort((a, b) => {
        const ta = new Date(a.date).getTime() || 0;
        const tb = new Date(b.date).getTime() || 0;
        return tb - ta;
      });
    }
    // Sort groups by Outlook's canonical order.
    const sortedGroups = Array.from(groups.entries()).sort(
      ([a], [b]) => dateGroupRank(a) - dateGroupRank(b)
    );
    return sortedGroups;
  }, [filteredMails]);

  return (
    <div className="border-x border-gray-200 flex flex-col h-full bg-white">
      <div className="p-4 border-b border-gray-200 flex justify-between items-center gap-2">
        <h2 className="m-0 text-xl font-semibold text-gray-800">{selectedFolderName}</h2>
        <FilterMenu />
      </div>
      <div className="flex-1 overflow-hidden">
        {isLoading ? (
          <div className="flex h-full items-center justify-center text-sm text-gray-500">
            Loading messages...
          </div>
        ) : filteredMails.length > 0 ? (
          // Render mails grouped by Outlook-style date groups.
          // We previously used a Syncfusion ListViewComponent but
          // rendering the grouped sections directly is simpler,
          // more flexible, and keeps the existing MailListItem
          // component unchanged. The grouped list scrolls in the
          // same outer container the ListView used.
          <div className="h-full overflow-y-auto custom-listview">
            {groupedMails.map(([groupLabel, mails]) => {
              const isCollapsed = collapsedGroups.has(groupLabel);
              return (
                <div key={groupLabel} className="mail-group">
                  {/* Sticky expand/collapse header. Clicking the
                      entire header toggles the group's collapsed
                      state — the chevron reflects that state and
                      the row count is shown on the right. */}
                  <button
                    type="button"
                    onClick={() => toggleGroup(groupLabel)}
                    className="sticky top-0 z-10 w-full flex items-center justify-between gap-2 bg-gray-50/95 backdrop-blur-sm px-4 py-1.5 text-xs font-semibold text-gray-600 uppercase tracking-wide border-b border-gray-200 hover:bg-gray-100 transition-colors text-left cursor-pointer border-x-0 border-t-0"
                    aria-expanded={!isCollapsed}
                    aria-label={isCollapsed ? `Expand ${groupLabel}` : `Collapse ${groupLabel}`}
                    data-mail-group={groupLabel}
                    data-collapsed={isCollapsed ? 'true' : 'false'}
                  >
                    <div className="flex items-center gap-1.5 min-w-0">
                      {isCollapsed ? (
                        <ChevronRight size={12} className="text-gray-400 flex-shrink-0" />
                      ) : (
                        <ChevronDown size={12} className="text-gray-400 flex-shrink-0" />
                      )}
                      <span className="truncate">{groupLabel}</span>
                    </div>
                    {/* Show the row count on the right when
                        collapsed so the user knows how many mails
                        are in the hidden section. Outlook shows
                        this same affordance. */}
                    {isCollapsed && (
                      <span className="text-[10px] font-normal text-gray-500 normal-case tracking-normal">
                        {mails.length} {mails.length === 1 ? 'item' : 'items'}
                      </span>
                    )}
                  </button>
                  {!isCollapsed && mails.map((mail) => (
                    <MailListItem
                      key={mail.id}
                      data={mail}
                      isSelected={mail.id === selectedMailId}
                      onSelect={handleMailSelect}
                      onFlagClick={handleFlagClick}
                      onReadClick={handleReadClick}
                    />
                  ))}
                </div>
              );
            })}
          </div>
        ) : (
          <div className="p-6 text-center text-gray-400 text-sm">
            {selectedFolderMails.length > 0 ? 'No matching results' : 'No items to display'}
          </div>
        )}
      </div>
    </div>
  );
};
