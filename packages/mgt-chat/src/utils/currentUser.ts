import { Providers } from '@microsoft/mgt-element';

const currentUserId = () => Providers.globalProvider.getActiveAccount?.().id.split('.')[0] || '';

export { currentUserId };
