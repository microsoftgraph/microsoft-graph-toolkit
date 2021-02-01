// NOTE : This is a simple cache plugin made for the purpose of demonstrating caching support for the Electron Provider.
// PLEASE DO NOT USE THIS IN PRODUCTION ENVIRONMENTS.

const fs = require('fs');
var path = require('path');

import { CACHE_LOCATION } from './Constants';

/**
 * Reads tokens from storage if exists and stores an in-memory copy.
 *
 * @param {*} cacheContext
 * @return {*}
 */
const beforeCacheAccess = async cacheContext => {
  console.log('PLEASE DO NOT USE THIS CACHE PLUGIN IN PRODUCTION ENVIRONMENTS!!!!');
  return new Promise<void>(async (resolve, reject) => {
    if (fs.existsSync(CACHE_LOCATION)) {
      fs.readFile(CACHE_LOCATION, 'utf-8', (err, data) => {
        if (err) {
          reject();
        } else {
          cacheContext.tokenCache.deserialize(data);
          resolve();
        }
      });
    } else {
      let dir = path.dirname(CACHE_LOCATION);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir);
      }
      fs.writeFile(CACHE_LOCATION, cacheContext.tokenCache.serialize(), err => {
        if (err) {
          reject();
        }
      });
    }
  });
};

/**
 *  Writes token to storage.
 *
 * @param {*} cacheContext
 */
const afterCacheAccess = async cacheContext => {
  if (cacheContext.cacheHasChanged) {
    let dir = path.dirname(CACHE_LOCATION);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    await fs.writeFile(CACHE_LOCATION, cacheContext.tokenCache.serialize(), err => {
      if (err) {
        console.log(err);
      }
    });
  }
};

// PLEASE DO NOT USE THIS IN PRODUCTION ENVIRONMENTS.
export const SimpleCachePlugin = {
  beforeCacheAccess,
  afterCacheAccess
};
