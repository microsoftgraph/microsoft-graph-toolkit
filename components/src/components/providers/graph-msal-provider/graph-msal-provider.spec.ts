jest.mock('@msgraphtoolkit/providers');

import { MsalProvider, LoginType, Providers} from '@msgraphtoolkit/providers'
import {MsalProviderComponent} from "./graph-msal-provider"

describe('MsalProviderComponent', () => {

  beforeEach(() => {
    jest.resetAllMocks();
  });

  it('clientId should be undefined by default', () => {
    const auth = new MsalProviderComponent();
    expect(auth.clientId).toEqual(undefined);
  });

  it('setting the clientId should init the MSALProvider with the clientId', async () => {
    const auth = new MsalProviderComponent();

    auth.clientId = "abc";
    auth.validateClientId();

    expect(Providers.add).toHaveBeenCalledTimes(1);
    expect(Providers.add).toBeCalledWith(expect.any(MsalProvider));
    expect(Providers.add).toBeCalledWith(expect.objectContaining({
      _clientId: "abc"
    }));
  });

  it('setting the loginType with an invalid type should not set the _loginType on the MSALProvider', async () => {
    const auth = new MsalProviderComponent();

    auth.clientId = "abc";
    auth.validateClientId();

    auth.loginType = "invalidLoginType";
    auth.validateLoginType();

    expect(Providers.add).toHaveBeenCalledTimes(2);
    expect(Providers.add).toBeCalledWith(expect.any(MsalProvider));
    expect(Providers.add).toBeCalledWith(expect.not.objectContaining({
      _loginType: "invalidLoginType"
    }));
  });

  it('setting the loginType with a valid type should set the _loginType on the MSALProvider', async () => {
    const auth = new MsalProviderComponent();

    auth.clientId = "abc";
    auth.validateClientId();

    auth.loginType = "popup";
    auth.validateLoginType();

    expect(Providers.add).toHaveBeenCalledTimes(2);
    expect(Providers.add).toBeCalledWith(expect.any(MsalProvider));
    expect(Providers.add).toBeCalledWith(expect.objectContaining({
      _loginType: LoginType.Popup
    }));
  });
});