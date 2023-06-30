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
    version: 2
  },
  users: {
    name: 'users',
    stores: {
      users: 'users',
      usersQuery: 'usersQuery',
      userFilters: 'userFilters'
    },
    version: 3
  },
  photos: {
    name: 'photos',
    stores: {
      contacts: 'contacts',
      users: 'users',
      groups: 'groups',
      teams: 'teams'
    },
    version: 2
  },
  people: {
    name: 'people',
    stores: {
      contacts: 'contacts',
      groupPeople: 'groupPeople',
      peopleQuery: 'peopleQuery'
    },
    version: 3
  },
  groups: {
    name: 'groups',
    stores: {
      groups: 'groups',
      groupsQuery: 'groupsQuery'
    },
    version: 5
  },
  get: {
    name: 'responses',
    stores: {
      responses: 'responses'
    },
    version: 2
  },
  search: {
    name: 'search',
    stores: {
      responses: 'responses'
    },
    version: 2
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
    version: 2
  },
  fileLists: {
    name: 'file-lists',
    stores: {
      fileLists: 'fileLists',
      insightfileLists: 'insightfileLists'
    },
    version: 2
  }
};
