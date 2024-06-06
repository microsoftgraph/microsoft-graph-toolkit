import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client';
export interface ExtendedAuthenticationProviderOptions extends AuthenticationProviderOptions {
  forceTokenRefresh?: boolean;
}
