import { IProvider } from './IProvider';
import { SimpleProvider } from './SimpleProvider';

describe('IProvider tests', () => {
  let p: IProvider;
  beforeEach(() => {
    p = new SimpleProvider(
      () => Promise.resolve('fake-token'),
      () => Promise.resolve(),
      () => Promise.resolve()
    );
    p.approvedScopes = ['user.read', 'group.read.all', 'presence.read'];
  });
  it('should provide an empty array when one scope is already present', () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.read.all']);
    expect(result).toEqual([]);
  });
  it('should provide an empty array when one scope is already present ignoring case of scopes in provider', () => {
    p.approvedScopes = ['user.read', 'Group.Read.All', 'presence.read'];
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.read.all']);
    expect(result).toEqual([]);
  });
  it('should provide an empty array when one scope is already present ignoring case of scopes in provider', () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'Group.Read.All']);
    expect(result).toEqual([]);
  });
  it('should provide an the first element in the pass array where there is no overlap', () => {
    const result = p.needsAdditionalScopes(['groupmember.read.all', 'group.readwrite.all']);
    expect(result).toEqual(['groupmember.read.all']);
  });
});
