jest.mock('../../auth/Auth');

import { MyMsalProvider } from './my-msal-provider';
import { Auth } from '../../auth/Auth'
import { MSALProvider } from '../../auth/MSALProvider';
import { LoginType } from '../../auth/IAuthProvider';

describe('MyMsalProvider', () => {

  beforeEach(() => {
    jest.resetAllMocks();
  });

  it('clientId should be undefined by default', () => {
    const auth = new MyMsalProvider();
    expect(auth.clientId).toEqual(undefined);
  });

  it('setting the clientId should init the MSALProvider with the clientId', async () => {
    const auth = new MyMsalProvider();

    auth.clientId = "abc";
    auth.validateClientId();

    expect(Auth.initWithProvider).toHaveBeenCalledTimes(1);
    expect(Auth.initWithProvider).toBeCalledWith(expect.any(MSALProvider));
    expect(Auth.initWithProvider).toBeCalledWith(expect.objectContaining({
      _clientId: "abc"
    }));
  });

  it('setting the loginType with an invalid type should not set the _loginType on the MSALProvider', async () => {
    const auth = new MyMsalProvider();

    auth.clientId = "abc";
    auth.validateClientId();

    auth.loginType = "invalidLoginType";
    auth.validateLoginType();

    expect(Auth.initWithProvider).toHaveBeenCalledTimes(2);
    expect(Auth.initWithProvider).toBeCalledWith(expect.any(MSALProvider));
    expect(Auth.initWithProvider).toBeCalledWith(expect.not.objectContaining({
      _loginType: "invalidLoginType"
    }));
  });

  it('setting the loginType with a valid type should set the _loginType on the MSALProvider', async () => {
    const auth = new MyMsalProvider();

    auth.clientId = "abc";
    auth.validateClientId();

    auth.loginType = "popup";
    auth.validateLoginType();

    expect(Auth.initWithProvider).toHaveBeenCalledTimes(2);
    expect(Auth.initWithProvider).toBeCalledWith(expect.any(MSALProvider));
    expect(Auth.initWithProvider).toBeCalledWith(expect.objectContaining({
      _loginType: LoginType.Popup
    }));
  });
});