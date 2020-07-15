import { MsalProvider, MsalConfig } from '../../src/providers/MsalProvider';
import { LoginType } from '../../src/providers/IProvider';
import { AuthenticationProviderOptions } from '@microsoft/microsoft-graph-client/lib/es/IAuthenticationProviderOptions';
import { UserAgentApplication } from 'msal';

jest.mock('@microsoft/microsoft-graph-client/lib/es/Client');
jest.mock('msal/lib-commonjs');

describe('MSALProvider', () => {
  beforeEach(() => {
    jest.resetAllMocks();
    UserAgentApplication.prototype.acquireTokenSilent = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        accessToken: 'abc'
      })
    );
  });

  it('undefined clientId should throw exception', () => {
    const config: MsalConfig = {
      clientId: undefined
    };
    expect(() => {
      new MsalProvider(config);
    }).toThrowError('clientId must be provided');
  });

  it('_loginType Popup should call the loginPopup', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    expect(msalProvider.provider.loginPopup).toBeCalled();
  });

  it('_loginType Redirect should call the loginRedirect', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };
    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    expect(msalProvider.provider.loginRedirect).toBeCalled();
  });

  it('getAccessToken should call acquireTokenSilent', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };
    const msalProvider = new MsalProvider(config);
    const authProviderOptions = {} as AuthenticationProviderOptions;
    const accessToken = await msalProvider.getAccessToken(authProviderOptions);
    expect(msalProvider.provider.acquireTokenSilent).toBeCalled();
  });

  it('logout should fire onLoginChanged callback', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);
    await msalProvider.login();
    expect.assertions(1);
    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.logout();
  });

  it('login should fire onLoginChanged callback when loginType is Popup', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Popup
    };

    UserAgentApplication.prototype.loginPopup = jest.fn().mockReturnValueOnce(
      Promise.resolve({
        account: 1
      })
    );

    const msalProvider = new MsalProvider(config);

    expect.assertions(1);
    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.login();
  });

  it.skip('login should fire onLoginChanged callback when loginType is Redirect', async () => {
    const config: MsalConfig = {
      clientId: 'abc',
      loginType: LoginType.Redirect
    };

    UserAgentApplication.prototype.loginRedirect = jest.fn().mockImplementationOnce(() => {
      //msalProvider.tokenReceivedCallback(undefined);
    });

    const msalProvider = new MsalProvider(config);

    expect.assertions(1);
    msalProvider.onStateChanged(() => {
      expect(true).toBeTruthy();
    });
    await msalProvider.login();
  });
});
