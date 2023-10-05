import { CacheOptions, CacheService } from '@microsoft/mgt-react';

export const isConversationCacheEnabled = (): boolean => {
  const conversation = CacheService.config.conversation as CacheOptions;
  return conversation.isEnabled && CacheService.config.isEnabled;
};
