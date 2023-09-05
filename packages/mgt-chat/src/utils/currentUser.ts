import { Providers } from '@microsoft/mgt-element';

const getCurrentUser = () => Providers.globalProvider.getActiveAccount?.();
const currentUserId = () => getCurrentUser()?.id.split('.')[0] || '';
const currentUserName = () => getCurrentUser()?.name || '';

export { getCurrentUser, currentUserId, currentUserName };
