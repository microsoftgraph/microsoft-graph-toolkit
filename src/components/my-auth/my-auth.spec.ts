import { MyAuth } from './my-auth';

describe('MyAuth', () => {
  it('clientId should be undefined by default', () => {
    const auth = new MyAuth();
    expect(auth.clientId).toEqual(undefined);
  });
});