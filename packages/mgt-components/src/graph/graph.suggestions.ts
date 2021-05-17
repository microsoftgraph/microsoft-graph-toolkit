/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { IGraph, prepScopes, CacheItem, CacheService, CacheStore } from '@microsoft/mgt-element';
import { schemas } from './cacheStores';

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionPeople extends CacheItem {
  /**
   * json representing a person stored as string
   */
  entity: string;
  displayName?: string;
  personImage?: string;
  givenName?: string;
  surname?: string;
  jobTitle?: string;
  referenceId: string;
  imAddress?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionFile extends CacheItem {
  /**
   * json representing a person stored as string
   */
  entity: string;
  name: string;
  addtionalInformation?: string;
  DateModified?: string;
  FileExtension?: string;
  FileType?: string;
  Text?: string;
  FileSize?: number;
  AccessUrl?: string;
  referenceId?: string;
  id?: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface SuggestionText extends CacheItem {
  /**
   * json representing a person stored as string
   */
  entity: string;
  text: string;
  referenceId: string;
}

/**
 * Object to be stored in cache representing individual people
 */
export interface Suggestions extends CacheItem {
  /**
   * json representing a person stored as string
   */
  fileSuggestions: SuggestionFile[];
  textSuggestions: SuggestionText[];
  peopleSuggestions: SuggestionPeople[];
}

export interface SuggestionQueryConfig extends CacheItem {
  maxFileCount: number;
  maxPeopleCount: number;
  maxTextCount: number;
  queryString: string;
  queryEntities: string;
  cvid?: string;
  textDecorations?: string;
}

/**
 * Stores results of queries (multiple people returned)
 */
interface CachePeopleQuery extends CacheItem {
  /**
   * max number of results the query asks for
   */
  maxResults?: number;
  /**
   * list of people returned by query (might be less than max results!)
   */
  results?: string[];
}

const getIsPeopleCacheEnabled = (): boolean =>
  CacheService.config.suggestions.isEnabled && CacheService.config.isEnabled;

/**
 * Defines the expiration time
 */
const getSuggestionInvalidationTime = (): number =>
  CacheService.config.suggestions.invalidationPeriod || CacheService.config.defaultInvalidationPeriod;

/**
 * async promise to the Graph for Suggestions, by default, it will request the most frequent contacts for the signed in user.
 *
 * @returns {(Promise<Person[]>)}
 * @memberof Graph
 */
export async function getSuggestions(graph: IGraph, queryConfig: SuggestionQueryConfig): Promise<Suggestions> {
  var mockFileSuggestions: SuggestionFile[] = [
    {
      entity: 'File',
      name: 'File A.docx',
      addtionalInformation: 'Append Description on Main Description Main Description Main Description Main Desc',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.1',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'docx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File B.xlsx',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.2',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'xlsx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File C.pptx',
      addtionalInformation: 'C Append Description on Main Description Main Description Main Description Main Desc',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.3',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'pptx',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    },
    {
      entity: 'File',
      name: 'File D.txt',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.4',
      FileSize: 947740,
      DateModified: '2019-07-19T00:54:58',
      FileExtension: 'txt',
      FileType: 'Link',
      AccessUrl:
        'https://microsofteur.sharepoint.com/teams/MicrosoftSearch/_layouts/15/Doc.aspx?sourcedoc=%7B44EB9A02-5F59-4CD5-89BD-47500F52E22D%7D&file=Vibranium-June2019.pptx&action=edit&mobileredirect=true&DefaultItemOpen=1'
    }
  ];

  var mockTextSuggestions: SuggestionText[] = [
    {
      entity: 'Text',
      text: 'Text A',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.5'
    },
    {
      entity: 'Text',
      text:
        'any one of the vector terms added to form a vector sum or resultant / a coordinate of a vectoreither member of an ordered pair of numbers',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.6'
    },
    {
      entity: 'Text',
      text: 'Text C',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.7'
    },
    {
      entity: 'Text',
      text: 'Text D',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.8'
    },
    {
      entity: 'Text',
      text: 'Text E',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10000.9'
    }
  ];

  var mockPeopleSuggestions: SuggestionPeople[] = [
    {
      entity: 'People',
      displayName: 'Isaiah Langer',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.10',
      imAddress: 'sip:isaiahl@m365x214355.onmicrosoft.com'
    },
    {
      entity: 'People',
      displayName: 'Megan Bowen',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.11',
      imAddress: 'sip:meganb@m365x214355.onmicrosoft.com'
    },
    {
      entity: 'People',
      displayName: 'Alex Wilber',
      jobTitle: 'Web Marketing Manager',
      //personImage: 'http://5b0988e595225.cdn.sohucs.com/images/20171130/658b9f5ea4394831b90e3f65d3cd83b6.jpeg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.12'
    },
    {
      entity: 'People',
      displayName: 'bob@tailspin.com',
      //personImage: 'http://image14.m1905.cn/uploadfile/2018/1107/20181107104420301408_watermark.jpg',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.13'
    },
    {
      entity: 'People',
      displayName: 'Lynne Robbins',
      personImage: '',
      referenceId: '36d3928a-7ccb-4359-9c17-fd9be6255364.10001.14',
      imAddress: 'sip:lynner@m365x214355.onmicrosoft.com'
    }
  ];

  var mockSuggestions: Suggestions = {
    fileSuggestions: mockFileSuggestions,
    textSuggestions: mockTextSuggestions,
    peopleSuggestions: mockPeopleSuggestions
  };
  if (queryConfig.queryEntities.toLowerCase().indexOf('file') != -1) {
    mockSuggestions.fileSuggestions = mockSuggestions.fileSuggestions
      .filter(file => {
        return file.name.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
      })
      .slice(0, queryConfig.maxFileCount);
  } else {
    mockSuggestions.fileSuggestions = [];
  }

  if (queryConfig.queryEntities.toLowerCase().indexOf('text') != -1) {
    mockSuggestions.textSuggestions = mockSuggestions.textSuggestions
      .filter(text => {
        return text.text.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
      })
      .slice(0, queryConfig.maxTextCount);
  } else {
    mockSuggestions.textSuggestions = [];
  }

  if (queryConfig.queryEntities.toLowerCase().indexOf('people') != -1) {
    mockSuggestions.peopleSuggestions = mockSuggestions.peopleSuggestions
      .filter(person => {
        return person.displayName.toLowerCase().startsWith(queryConfig.queryString.toLowerCase());
      })
      .slice(0, queryConfig.maxPeopleCount);
  } else {
    mockSuggestions.peopleSuggestions = [];
  }

  return mockSuggestions;
}
