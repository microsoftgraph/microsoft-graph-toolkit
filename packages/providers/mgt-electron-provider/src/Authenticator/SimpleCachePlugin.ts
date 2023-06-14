/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-misused-promises */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
/* eslint-disable @typescript-eslint/no-var-requires */

/**
 * NOTE : This is a simple cache plugin made for the purpose of demonstrating caching support for the Electron Provider.
 * PLEASE DO NOT USE THIS IN PRODUCTION ENVIRONMENTS.
 */

const fs = require('fs');
const path = require('path');

import { CACHE_LOCATION } from './Constants';

/**
 * Reads tokens from storage if exists and stores an in-memory copy.
 *
 * @param {*} cacheContext
 * @return {*}
 */
const beforeCacheAccess = async cacheContext => {
  // eslint-disable-next-line no-console
  console.warn('🦒: PLEASE DO NOT USE THIS CACHE PLUGIN IN PRODUCTION ENVIRONMENTS!!!!');
  return new Promise<void>((resolve, reject) => {
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
      const dir = path.dirname(CACHE_LOCATION);
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
 * Writes token to storage.
 *
 * @param {*} cacheContext
 */
const afterCacheAccess = async cacheContext => {
  if (cacheContext.cacheHasChanged) {
    const dir = path.dirname(CACHE_LOCATION);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir);
    }
    await fs.writeFile(CACHE_LOCATION, cacheContext.tokenCache.serialize(), err => {
      if (err) {
        // eslint-disable-next-line no-console
        console.log('🦒: ', err);
      }
    });
  }
};

/**
 * PLEASE DO NOT USE THIS IN PRODUCTION ENVIRONMENTS.
 */
// eslint-disable-next-line @typescript-eslint/naming-convention
export const SimpleCachePlugin = {
  beforeCacheAccess,
  afterCacheAccess
};
