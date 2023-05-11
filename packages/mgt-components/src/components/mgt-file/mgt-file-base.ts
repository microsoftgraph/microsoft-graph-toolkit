import { property } from 'lit/decorators.js';
import { MgtTemplatedComponent } from '@microsoft/mgt-element';
import { OfficeGraphInsightString } from '../../graph/types';

/**
 * Provides base property definitions for use in file and file list components
 *
 * @abstract
 * @class MgtFileBase
 * @extends {MgtTemplatedComponent}
 */
export abstract class MgtFileBase extends MgtTemplatedComponent {
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _listId: string;
  private _groupId: string;
  private _userId: string;
  private _insightType: OfficeGraphInsightString;
  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'site-id'
  })
  public get siteId(): string {
    return this._siteId;
  }
  public set siteId(value: string) {
    if (value === this._siteId) {
      return;
    }

    this._siteId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide drive id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'drive-id'
  })
  public get driveId(): string {
    return this._driveId;
  }
  public set driveId(value: string) {
    if (value === this._driveId) {
      return;
    }

    this._driveId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide group id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'group-id'
  })
  public get groupId(): string {
    return this._groupId;
  }
  public set groupId(value: string) {
    if (value === this._groupId) {
      return;
    }

    this._groupId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide list id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'list-id'
  })
  public get listId(): string {
    return this._listId;
  }
  public set listId(value: string) {
    if (value === this._listId) {
      return;
    }

    this._listId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide user id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'user-id'
  })
  public get userId(): string {
    return this._userId;
  }
  public set userId(value: string) {
    if (value === this._userId) {
      return;
    }

    this._userId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'item-id'
  })
  public get itemId(): string {
    return this._itemId;
  }
  public set itemId(value: string) {
    if (value === this._itemId) {
      return;
    }

    this._itemId = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   * @memberof MgtFile
   */
  @property({
    attribute: 'item-path'
  })
  public get itemPath(): string {
    return this._itemPath;
  }
  public set itemPath(value: string) {
    if (value === this._itemPath) {
      return;
    }

    this._itemPath = value;
    void this.requestStateUpdate();
  }

  /**
   * allows developer to provide insight type for a file
   * can be trending, used, or shared
   *
   * @type {OfficeGraphInsightString}
   * @memberof MgtFile
   */
  @property({
    attribute: 'insight-type'
  })
  public get insightType(): OfficeGraphInsightString {
    return this._insightType;
  }
  public set insightType(value: OfficeGraphInsightString) {
    if (value === this._insightType) {
      return;
    }

    this._insightType = value;
    void this.requestStateUpdate();
  }
}
