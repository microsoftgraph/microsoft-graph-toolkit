import { replaceOrAppendSessionId } from './replaceOrAppendSessionId';
import { expect } from '@open-wc/testing';

describe('replaceOrAppendSessionId tests', () => {
  it('should replace existing sessionid', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me?sessionid=default', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?sessionid=123');
  });
  it('should replace existing sessionid when there are query string params after sessionid', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me?sessionid=default&foo=bar', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?sessionid=123&foo=bar');
  });
  it('should replace existing sessionid when there are query string params before sessionid', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me?foo=bar&sessionid=default', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?foo=bar&sessionid=123');
  });
  it('should append sessionid there is a query string params with sessionid as a substring', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me?foosessionid=bar', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?foosessionid=bar&sessionid=123');
  });
  it('should replace sessionid there is also a query string params with sessionid as a substring', async () => {
    const result = replaceOrAppendSessionId(
      'https://graph.microsoft.com/1.0/me?foosessionid=bar&sessionid=default&bar=too',
      '123'
    );
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?foosessionid=bar&sessionid=123&bar=too');
  });
  it('should append sessionid with ? when no query params present', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?sessionid=123');
  });
  it('should append sessionid with & when query params are present', async () => {
    const result = replaceOrAppendSessionId('https://graph.microsoft.com/1.0/me?foo=bar', '123');
    await expect(result).to.equal('https://graph.microsoft.com/1.0/me?foo=bar&sessionid=123');
  });
});
