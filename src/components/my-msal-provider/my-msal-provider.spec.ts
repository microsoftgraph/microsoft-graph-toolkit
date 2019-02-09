import { MyMsalProvider } from './my-msal-provider';

describe('MyMsalProvider', () => {
  it('clientId should be undefined by default', () => {
    const auth = new MyMsalProvider();
    expect(auth.clientId).toEqual(undefined);
  });
});