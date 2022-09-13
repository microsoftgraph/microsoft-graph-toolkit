/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

/**
 * All schemas and stores for caching component calls
 */
export const schemas = {
  presence: {
    name: 'presence',
    stores: {
      presence: 'presence'
    },
    version: 1
  },
  users: {
    name: 'users',
    stores: {
      users: 'users',
      usersQuery: 'usersQuery',
      userFilters: 'userFilters'
    },
    version: 2
  },
  photos: {
    name: 'photos',
    stores: {
      contacts: 'contacts',
      users: 'users',
      groups: 'groups'
    },
    version: 1
  },
  people: {
    name: 'people',
    stores: {
      contacts: 'contacts',
      groupPeople: 'groupPeople',
      peopleQuery: 'peopleQuery'
    },
    version: 2
  },
  groups: {
    name: 'groups',
    stores: {
      groups: 'groups',
      groupsQuery: 'groupsQuery'
    },
    version: 4
  },
  get: {
    name: 'responses',
    stores: {
      responses: 'responses'
    },
    version: 1
  },
  files: {
    name: 'files',
    stores: {
      driveFiles: 'driveFiles',
      groupFiles: 'groupFiles',
      siteFiles: 'siteFiles',
      userFiles: 'userFiles',
      insightFiles: 'insightFiles',
      fileQueries: 'fileQueries'
    },
    version: 1
  },
  fileLists: {
    name: 'file-lists',
    stores: {
      fileLists: 'fileLists',
      insightfileLists: 'insightfileLists'
    },
    version: 1
  }
};
