import { SPHttpClient } from '@microsoft/sp-http';

export interface IFloatingFeedbackProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userEmail: string;
  listName: string; // Changed from listId
  pageName: string;
  position: 'Top' | 'Bottom';
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
