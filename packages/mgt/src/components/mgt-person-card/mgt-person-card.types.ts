import { Message, Person, SharedInsight, User } from '@microsoft/microsoft-graph-types';
import { Profile } from '@microsoft/microsoft-graph-types-beta';
import { IDynamicPerson } from '../../graph/types';

export interface MgtPersonCardConfig {
  /**
   * Sets or gets whether the person card component can use Contacts APIs to
   * find contacts and their images
   *
   * @type {boolean}
   */
  useContactApis: boolean;
  isSendMessageVisible: boolean;
  sections: {
    contact: boolean;
    organization: boolean;
    mailMessages: boolean;
    files: boolean;
    profile: boolean;
  };
}

export interface MgtPersonCardHistoryState {
  state: MgtPersonCardState;
  personDetails: IDynamicPerson;
  personImage: string;
}

export type UserWithManager = User & { manager?: UserWithManager };

export interface MgtPersonCardState {
  directReports?: User[];
  files?: SharedInsight[];
  messages?: Message[];
  people?: Person[];
  person?: UserWithManager;
  profile?: Profile;
}
