import { SPHttpClient } from '@microsoft/sp-http';

export interface IClientAzDirectoryProps {
  listTitle: string;
  nameField: string;
  secondaryField: string;
  linkField: string;
  spHttpClient: SPHttpClient;
  webAbsoluteUrl: string;
  hasTeamsContext: boolean;
}
