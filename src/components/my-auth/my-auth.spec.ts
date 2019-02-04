import { MyAuth } from './my-auth';

describe('MyAuth', () => {
  it('name should be undefined by default', () => {
    const auth = new MyAuth();

    expect(auth.name).toEqual(undefined);
  });
});