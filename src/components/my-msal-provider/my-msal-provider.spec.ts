import { MyMsalProvider } from './my-msal-provider';
import * as Auth from '../../auth/Auth'
import { MSALProvider } from '../../auth/MSALProvider';

describe('MyMsalProvider', () => {
  it('clientId should be undefined by default', () => {
    const auth = new MyMsalProvider();
    expect(auth.clientId).toEqual(undefined);
  });

  it('setting the clientId should init the MSALProvider', async () => {
    const auth = new MyMsalProvider();

    const spy = jest.spyOn(Auth, "initWithProvider");

    auth.clientId = "abc";
    auth.validateClientId();

    expect(spy).toHaveBeenCalledTimes(1);
    expect(spy).toBeCalledWith(expect.any(MSALProvider));
    expect(spy).toBeCalledWith(expect.objectContaining({
      _clientId: "abc"
    }));
  });
});