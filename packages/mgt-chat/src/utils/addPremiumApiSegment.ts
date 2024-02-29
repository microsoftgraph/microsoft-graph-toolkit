import { GraphConfig } from '../statefulClient/GraphConfig';

export const addPremiumApiSegment = (url: string) => {
  // early exit if premium apis are not enabled
  if (!GraphConfig.usePremiumApis) {
    return url;
  }
  const urlHasExistingQueryParams = url.includes('?');
  return `${url}${urlHasExistingQueryParams ? '&' : '?'}model=B`;
};
