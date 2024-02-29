import { GraphConfig } from '../statefulClient/GraphConfig';
import { addPremiumApiSegment } from './addPremiumApiSegment';
import { expect } from '@open-wc/testing';

describe('addPremiumApiSegment tests', () => {
  it('should return the original url if premium apis are not enabled', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me';
    const result = addPremiumApiSegment(url);
    await expect(result).to.eql(url);
  });

  it('should add the premium api segment to the url', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me';
    GraphConfig.usePremiumApis = true;
    const result = addPremiumApiSegment(url);
    await expect(result).to.eql(`${url}?model=B`);
  });

  it('should add the premium api segment to the url when it already has query params', async () => {
    const url = 'https://graph.microsoft.com/v1.0/me?select=id';
    GraphConfig.usePremiumApis = true;
    const result = addPremiumApiSegment(url);
    await expect(result).to.eql(`${url}&model=B`);
  });
});
