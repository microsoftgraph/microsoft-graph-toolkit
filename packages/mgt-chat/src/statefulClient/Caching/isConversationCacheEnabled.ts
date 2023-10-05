import { CacheService } from '@microsoft/mgt-react';

export const isConversationCacheEnabled = (): boolean => {
  // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
  const conversation = CacheService.config.conversation;
  // eslint-disable-next-line @typescript-eslint/no-unsafe-return, @typescript-eslint/no-unsafe-member-access
  return conversation.isEnabled && CacheService.config.isEnabled;
};
