import { CacheService } from '@microsoft/mgt-react';

export const isConversationCacheEnabled = (): boolean =>
  CacheService.config.conversation.isEnabled && CacheService.config.isEnabled;
