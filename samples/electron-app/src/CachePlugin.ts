/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const fs = require('fs');

import { CACHE_LOCATION } from './Constants';

const beforeCacheAccess = async cacheContext => {
  return new Promise<void>(async (resolve, reject) => {
    if (fs.existsSync(CACHE_LOCATION)) {
      console.log('Cache exists');
      fs.readFile(CACHE_LOCATION, 'utf-8', (err, data) => {
        if (err) {
          reject();
        } else {
          try {
            cacheContext.tokenCache.deserialize(data);
          } catch (e) {
            fs.writeFile(CACHE_LOCATION, '', err => {
              reject();
            });
          }
          resolve();
        }
      });
    } else {
      console.log('Cache does not exist');
      const str = cacheContext.tokenCache.serialize();
      fs.writeFile(CACHE_LOCATION, str, err => {
        if (err) {
          reject();
        }
      });
    }
  });
};

const afterCacheAccess = async cacheContext => {
  if (cacheContext.cacheHasChanged) {
    const str = cacheContext.tokenCache.serialize();
    await fs.writeFile(CACHE_LOCATION, str, err => {
      if (err) {
        console.log(err);
      }
    });
  }
};

export const cachePlugin = {
  beforeCacheAccess,
  afterCacheAccess
};
