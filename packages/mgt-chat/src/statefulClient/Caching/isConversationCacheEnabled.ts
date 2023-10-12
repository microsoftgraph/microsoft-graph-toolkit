import { CacheService } from '@microsoft/mgt-react';

export const isConversationCacheEnabled = (): boolean => {
  const conversation = CacheService.config.conversation;
  return conversation.isEnabled && CacheService.config.isEnabled;
};
