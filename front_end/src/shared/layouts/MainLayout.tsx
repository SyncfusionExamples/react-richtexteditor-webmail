import React from 'react';
import { SplitterComponent, PanesDirective, PaneDirective } from '@syncfusion/ej2-react-layouts';
import { HeaderToolbar } from '../components/HeaderToolbar';
import { FolderNavigation } from '../../modules/folders/FolderNavigation';
import { MailList } from '../../modules/mail/MailList';
import { MailReader } from '../../modules/mail/MailReader';
import { ComposeEditor } from '../../modules/compose/ComposeEditor';
import { useUIStore } from '../../store/Store';
import { useAuth } from '../../auth/AuthProvider';

export const MainLayout: React.FC = () => {
  const { isComposing } = useUIStore();
  const auth = useAuth();

  if (auth.isLoadingData) {
    return (
      <div className="flex items-center justify-center h-screen w-screen bg-white">
        <div className="text-center px-6 py-4 rounded-xl border border-gray-200 shadow-sm bg-gray-50">
          <div className="text-lg font-semibold text-gray-900">Loading mailbox…</div>
          <div className="text-sm text-gray-600 mt-2">Please wait while your mail content loads.</div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col h-screen w-screen bg-white">
      <HeaderToolbar />
      <div className="flex-1 overflow-hidden">
        <SplitterComponent id="main-splitter" width="100%" height="100%" separatorSize={4}>
          <PanesDirective>
            <PaneDirective size="15%" min="12%" content={() => <FolderNavigation />} />
            <PaneDirective size="35%" min="20%" content={() => <MailList />} />
            <PaneDirective size="50%" min="30%" content={() => isComposing ? <ComposeEditor /> : <MailReader />} />
          </PanesDirective>
        </SplitterComponent>
      </div>
    </div>
  );
};
