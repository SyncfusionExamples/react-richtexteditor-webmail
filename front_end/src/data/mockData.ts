export interface FolderData {
  [key: string]: any;
  id: string;
  name: string;
  icon?: string;
  hasChild?: boolean;
  expanded?: boolean;
  parentId?: string;
  count?: number;
}

export const folderData: FolderData[] = [
  // Favorites
  { id: '1', name: 'Favorites', hasChild: true, expanded: true },
  { id: '1-1', parentId: '1', name: 'Inbox', icon: 'Inbox', count: 12 },
  { id: '1-2', parentId: '1', name: 'Sent Items', icon: 'Send' },

  // Main Folders
  { id: '2', name: 'Folders', hasChild: true, expanded: true },
  { id: '2-1', parentId: '2', name: 'Inbox', icon: 'Inbox' },
  { id: '2-2', parentId: '2', name: 'Drafts', icon: 'FileText' },
  { id: '2-3', parentId: '2', name: 'Sent Items', icon: 'Send' },
  { id: '2-4', parentId: '2', name: 'Deleted Items', icon: 'Trash2' },
  { id: '2-5', parentId: '2', name: 'Archive', icon: 'Archive', hasChild: true, expanded: false },
  { id: '2-6', parentId: '2', name: 'Junk Email', icon: 'AlertCircle' },

  // Archive - Year based nested structure
  { id: '2-5-1', parentId: '2-5', name: '2026', hasChild: true, expanded: false },
  { id: '2-5-1-1', parentId: '2-5-1', name: 'January' },
  { id: '2-5-1-2', parentId: '2-5-1', name: 'February' },
  { id: '2-5-1-3', parentId: '2-5-1', name: 'March' },
  { id: '2-5-1-4', parentId: '2-5-1', name: 'April' },

  { id: '2-5-2', parentId: '2-5', name: '2025', hasChild: true, expanded: false },
  { id: '2-5-2-1', parentId: '2-5-2', name: 'Q4 (Oct-Dec)' },
  { id: '2-5-2-2', parentId: '2-5-2', name: 'Q3 (Jul-Sep)' },

  // Custom Folders - Work Projects
  { id: '3', name: 'Work Projects', hasChild: true, expanded: false },
  { id: '3-1', parentId: '3', name: 'Clients', hasChild: true, expanded: false },
  { id: '3-1-1', parentId: '3-1', name: 'Acme Corp' },
  { id: '3-1-2', parentId: '3-1', name: 'TechStart Inc' },
  { id: '3-1-3', parentId: '3-1', name: 'Global Solutions' },
  { id: '3-1-4', parentId: '3-1', name: 'Blue Wave LLC' },

  { id: '3-2', parentId: '3', name: 'Internal Projects', hasChild: true, expanded: false },
  { id: '3-2-1', parentId: '3-2', name: 'Website Redesign' },
  { id: '3-2-2', parentId: '3-2', name: 'Mobile App Launch' },
  { id: '3-2-3', parentId: '3-2', name: 'API Development' },

  { id: '3-3', parentId: '3', name: 'Meetings & Calls' },

  // Custom Folders - Personal
  { id: '4', name: 'Personal', hasChild: true, expanded: false },
  { id: '4-1', parentId: '4', name: 'Family' },
  { id: '4-2', parentId: '4', name: 'Friends' },
  { id: '4-3', parentId: '4', name: 'Subscriptions', hasChild: true, expanded: false },
  { id: '4-3-1', parentId: '4-3', name: 'Newsletters' },
  { id: '4-3-2', parentId: '4-3', name: 'E-Commerce' },

  // Custom Folders - Finance
  { id: '5', name: 'Finance & Billing', hasChild: true, expanded: false },
  { id: '5-1', parentId: '5', name: 'Invoices' },
  { id: '5-2', parentId: '5', name: 'Receipts' },
  { id: '5-3', parentId: '5', name: 'Bank Statements' },
  { id: '5-4', parentId: '5', name: 'Insurance' },

  // Custom Folders - Documentation
  { id: '6', name: 'Documentation', hasChild: true, expanded: false },
  { id: '6-1', parentId: '6', name: 'Contracts' },
  { id: '6-2', parentId: '6', name: 'Legal Documents' },
  { id: '6-3', parentId: '6', name: 'Certifications' },
];

export interface ContactData {
  [key: string]: any;
  Name: string;
  Email: string;
  Eimg?: string;
  Color?: string;
}

export const contactsData: ContactData[] = [
  { Name: 'Alex Johnson', Email: 'alex.j@example.com', Eimg: '1', Color: '#0078d4' },
  { Name: 'Marketing Team', Email: 'marketing@example.com', Eimg: '2', Color: '#107c10' },
  { Name: 'Sarah Smith', Email: 'sarah.smith@example.com', Eimg: '3', Color: '#d13438' },
  { Name: 'Robert Downey', Email: 'robert.d@example.com', Eimg: '4', Color: '#4c4a48' },
  { Name: 'John Wick', Email: 'john.w@example.com', Eimg: '5', Color: '#0078d4' },
  { Name: 'Peter Parker', Email: 'peter.p@example.com', Eimg: '6', Color: '#107c10' },
  { Name: 'Bruce Wayne', Email: 'bruce.w@example.com', Eimg: '7', Color: '#d13438' },
];

/**
 * Metadata for a single attachment on a mail message, as returned by
 * Microsoft Graph. We use this shape both for the list view and for
 * the inline display in the mail reader.
 *
 * `contentBytes` is intentionally NOT included here — attachments
 * are downloaded on demand from the backend's
 * `/mail/messages/:id/attachments/:attachmentId` endpoint so the
 * initial mail load doesn't carry large base64 payloads.
 */
export interface MailAttachment {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline?: boolean;
  lastModifiedDateTime?: string;
}

export interface MailData {
  [key: string]: any;
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
  /**
   * True when the message has at least one attachment (used as a
   * quick UI hint before attachments are loaded).
   */
  hasAttachments?: boolean;
  /**
   * Detailed attachment metadata. Populated by `getMessageDetails`
   * in the store when a mail is opened, NOT by the list endpoint.
   */
  attachments?: MailAttachment[];
}

export const mailData: MailData[] = [
  // Inbox mails
  {
    id: 'm1',
    folderId: '2-1',
    senderName: 'Alex Johnson',
    senderEmail: 'alex.j@example.com',
    subject: 'Project Update: Q3 Goals',
    snippet: 'Hi team, please find the updated Q3 goals attached. We have made some revisions based on...',
    date: '10:30 AM',
    isRead: false,
    isStarred: true,
    isFlagged: false,
    avatarColor: '#0078d4'
  },
  {
    id: 'm2',
    folderId: '2-1',
    senderName: 'Marketing Team',
    senderEmail: 'marketing@example.com',
    subject: 'New Campaign Launch',
    snippet: 'The new summer campaign is ready to go. Please review the final assets before we push to production.',
    date: 'Yesterday',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#107c10'
  },
  {
    id: 'm3',
    folderId: '2-1',
    senderName: 'Sarah Smith',
    senderEmail: 'sarah.smith@example.com',
    subject: 'Lunch next week?',
    snippet: 'Are you free for lunch next Tuesday? Would love to catch up and discuss the new project.',
    date: 'Oct 12',
    isRead: false,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#d13438'
  },
  {
    id: 'm5',
    folderId: '2-1',
    senderName: 'Robert Downey',
    senderEmail: 'robert.d@example.com',
    subject: 'Budget Review Meeting',
    snippet: 'Please prepare the budget report for the quarterly review. Meeting scheduled for Thursday at 2 PM.',
    date: 'Oct 11',
    isRead: true,
    isStarred: false,
    isFlagged: true,
    avatarColor: '#4c4a48'
  },
  {
    id: 'm6',
    folderId: '2-1',
    senderName: 'John Wick',
    senderEmail: 'john.w@example.com',
    subject: 'System Maintenance Alert',
    snippet: 'The server maintenance is scheduled for Saturday 2-4 AM. Expected downtime: 30 minutes.',
    date: 'Oct 10',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },

  // Sent Items
  {
    id: 'm7',
    folderId: '2-3',
    senderName: 'Me',
    senderEmail: 'me@example.com',
    subject: 'Re: Project Update: Q3 Goals',
    snippet: 'Thanks Alex. I will review these today and get back to you with my thoughts.',
    date: '11:00 AM',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },
  {
    id: 'm8',
    folderId: '2-3',
    senderName: 'Me',
    senderEmail: 'me@example.com',
    subject: 'Campaign Assets Approved',
    snippet: 'I have reviewed all the campaign assets and they look great. Ready to proceed with launch next week.',
    date: 'Oct 11',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },

  // Drafts
  {
    id: 'm9',
    folderId: '2-2',
    senderName: 'Me',
    senderEmail: 'me@example.com',
    subject: 'Draft: Team Building Event Planning',
    snippet: 'Hi everyone, I wanted to get your feedback on potential dates for our team building event...',
    date: 'Oct 9',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },
  {
    id: 'm10',
    folderId: '2-2',
    senderName: 'Me',
    senderEmail: 'me@example.com',
    subject: 'Draft: Feedback on Proposal',
    snippet: 'I have some initial thoughts on the proposal. I think we need to refine the approach...',
    date: 'Oct 8',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },

  // Work Projects - Clients
  {
    id: 'm11',
    folderId: '3-1-1',
    senderName: 'John Smith',
    senderEmail: 'john.smith@acmecorp.com',
    subject: 'Contract Renewal Discussion',
    snippet: 'We would like to discuss the renewal terms for our current contract. Are you available next week?',
    date: 'Oct 12',
    isRead: false,
    isStarred: false,
    isFlagged: true,
    avatarColor: '#107c10'
  },
  {
    id: 'm12',
    folderId: '3-1-2',
    senderName: 'Emily Chen',
    senderEmail: 'emily@techstart.io',
    subject: 'Feature Request - Mobile App',
    snippet: 'We have identified several features that would enhance user engagement. Let me share the details...',
    date: 'Oct 10',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#d13438'
  },

  // Work Projects - Internal
  {
    id: 'm13',
    folderId: '3-2-1',
    senderName: 'Design Team',
    senderEmail: 'design@company.com',
    subject: 'Website Redesign: Initial Mockups Ready',
    snippet: 'The initial mockups for the website redesign are ready for review. Please check the design portal...',
    date: 'Oct 11',
    isRead: false,
    isStarred: true,
    isFlagged: false,
    avatarColor: '#4c4a48'
  },
  {
    id: 'm14',
    folderId: '3-2-2',
    senderName: 'Dev Team',
    senderEmail: 'dev@company.com',
    subject: 'Mobile App: Development Sprint Update',
    snippet: 'Sprint 15 is on track. We have completed 85% of the planned features. Next: QA testing phase.',
    date: 'Oct 10',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#107c10'
  },

  // Personal - Family
  {
    id: 'm15',
    folderId: '4-1',
    senderName: 'Mom',
    senderEmail: 'mom@email.com',
    subject: 'Thanksgiving Plans?',
    snippet: 'Hi! Are you planning to come home for Thanksgiving? Let me know so I can prepare.',
    date: 'Oct 12',
    isRead: false,
    isStarred: true,
    isFlagged: false,
    avatarColor: '#d13438'
  },

  // Finance - Invoices
  {
    id: 'm16',
    folderId: '5-1',
    senderName: 'Stripe Billing',
    senderEmail: 'billing@stripe.com',
    subject: 'Invoice #INV-2026-0847',
    snippet: 'Your monthly invoice for April 2026 is ready. Amount due: $1,250.00',
    date: 'Oct 8',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#0078d4'
  },

  // Finance - Receipts
  {
    id: 'm17',
    folderId: '5-2',
    senderName: 'Amazon',
    senderEmail: 'order-confirmation@amazon.com',
    subject: 'Order Confirmation: Monitor Stand',
    snippet: 'Your order has been confirmed. Tracking number: 123456789. Estimated delivery: Oct 18.',
    date: 'Oct 6',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#FF9900'
  },

  // Documentation - Contracts
  {
    id: 'm18',
    folderId: '6-1',
    senderName: 'Legal Team',
    senderEmail: 'legal@company.com',
    subject: 'Contract Review: Vendor Agreement',
    snippet: 'Please review the attached vendor agreement and provide your comments by end of week.',
    date: 'Oct 9',
    isRead: false,
    isStarred: false,
    isFlagged: true,
    avatarColor: '#4c4a48'
  },

  // Archived mails
  {
    id: 'm19',
    folderId: '2-5-1-1',
    senderName: 'Old Project Team',
    senderEmail: 'project@archive.com',
    subject: 'Project Closure: Year-End Summary',
    snippet: 'Archiving all project communications from January. Final report attached.',
    date: 'Jan 31',
    isRead: true,
    isStarred: false,
    isFlagged: false,
    avatarColor: '#107c10'
  }
];
