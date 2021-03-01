import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { customElement, html, property, TemplateResult } from 'lit-element';
import { styles } from './mgt-file-picker-css';
import { arraysAreEqual, MgtTemplatedComponent, Providers, ProviderState } from '@microsoft/mgt-element';
import { OfficeGraphInsightString } from '../../graph/types';
import { MgtFlyout } from '../sub-components/mgt-flyout/mgt-flyout';
import { classMap } from 'lit-html/directives/class-map';
import { MgtFileList } from '../mgt-file-list/mgt-file-list';
import { FluentDesignSystemProvider, FluentButton } from '@fluentui/web-components';
import {
  getFilesByListQuery,
  getFilesByQueries,
  getFilesById,
  getFilesByPath,
  getMyInsightsFiles,
  getFiles,
  getDriveFilesById,
  getDriveFilesByPath,
  getGroupFilesById,
  getGroupFilesByPath,
  getSiteFilesById,
  getSiteFilesByPath
} from '../../graph/graph.files';

// Prevent tree-shaking
// FluentButton;
// FluentDesignSystemProvider;

/**
 * The File component is used to represent an individual file/folder from OneDrive or SharePoint by displaying information such as the file/folder name, an icon indicating the file type, and other properties such as the author, last modified date, or other details selected by the developer.
 *
 * @export
 * @class MgtFilePicker
 * @extends {MgtTemplatedComponent}
 *
 * @fires fileSelected
 *
 * @cssprop
 */

@customElement('mgt-file-picker')
export class MgtFilePicker extends MgtTemplatedComponent {
  /**
   * Array of styles to apply to the element. The styles should be defined
   * using the `css` tag function.
   */
  static get styles() {
    return styles;
  }

  /**
   * allows developer to provide query for a file list
   *
   * @type {string}
   * @memberof MgtFilePicker
   */
  @property({
    attribute: 'file-list-query'
  })
  public get fileListQuery(): string {
    return this._fileListQuery;
  }
  public set fileListQuery(value: string) {
    if (value === this._fileListQuery) {
      return;
    }

    this._fileListQuery = value;
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide an array of file queries
   *
   * @type {string[]}
   * @memberof MgtFilePicker
   */
  @property({
    attribute: 'file-queries',
    converter: (value, type) => {
      return value.split(',').map(v => v.trim());
    }
  })
  public get fileQueries(): string[] {
    return this._fileQueries;
  }
  public set fileQueries(value: string[]) {
    if (arraysAreEqual(this._fileQueries, value)) {
      return;
    }

    this._fileQueries = value;
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide an array of files
   *
   * @type {MicrosoftGraph.DriveItem[]}
   * @memberof MgtFilePicker
   */
  @property({ type: Object })
  public get files(): MicrosoftGraph.DriveItem[] {
    return this._files;
  }
  public set files(value: MicrosoftGraph.DriveItem[]) {
    if (value === this._files) {
      return;
    }

    this._files = value;
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide site id for a file
   *
   * @type {string}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide drive id for a file
   *
   * @type {string}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide group id for a file
   *
   * @type {string}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide item id for a file
   *
   * @type {string}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide item path for a file
   *
   * @type {string}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * allows developer to provide insight type for a file
   * can be trending, used, or shared
   *
   * @type {OfficeGraphInsightString}
   * @memberof MgtFilePicker
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
    this.requestStateUpdate();
  }

  /**
   * A number value to indicate the maximum number of files to show
   * @type {number}
   * @memberof MgtFilePicker
   */
  @property({
    attribute: 'show-max',
    type: Number
  })
  public showMax: number;

  /**
   * Gets the flyout element
   *
   * @protected
   * @type {MgtFlyout}
   * @memberof MgtFilePicker
   */
  protected get flyout(): MgtFlyout {
    return this.renderRoot.querySelector('.flyout');
  }

  /**
   * The selected item
   *
   * @readonly
   * @type {MicrosoftGraph.DriveItem}
   * @memberof MgtFilePicker
   */
  public get selectedItem() {
    return this._selectedItem;
  }

  private _fileListQuery: string;
  private _fileQueries: string[];
  private _files: MicrosoftGraph.DriveItem[];
  private _siteId: string;
  private _itemId: string;
  private _driveId: string;
  private _itemPath: string;
  private _groupId: string;
  private _insightType: OfficeGraphInsightString;
  private _selectedItem: MicrosoftGraph.DriveItem;
  private _doLoad: boolean;

  constructor() {
    super();

    this.showMax = 10;
  }

  public render() {
    if (!this.files && this.isLoadingState) {
      return this.renderLoading();
    }

    if (!this.files || this.files.length === 0) {
      return this.renderNoData();
    }

    // return this.renderTemplate('default', { files: this.files, max: this.showMax }) || this.renderFileList();
    const flyoutClasses = classMap({
      flyout: true,
      disabled: !Providers.globalProvider || Providers.globalProvider.state === ProviderState.SignedOut,
      loading: this.isLoadingState
    });

    return html`
      <mgt-flyout class=${flyoutClasses} light-dismiss>
        ${this.renderButton()} ${this.renderFlyout()}
      </mgt-flyout>
    `;
  }

  /**
   * Render the loading state
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderLoading(): TemplateResult {
    return this.renderTemplate('loading', null) || html``;
  }

  /**
   * Render the state when no data is available
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderNoData(): TemplateResult {
    return this.renderTemplate('no-data', null) || html``;
  }

  /**
   * Render the file list.
   *
   * @protected
   * @param {*} files
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderFlyout(): TemplateResult {
    return html`
      <div slot="flyout">
        <mgt-file-list
          id="file-list"
          .files=${this.files}
          .insightType=${this.insightType}
          .fileListQuery=${this.fileListQuery}
          .fileQueries=${this.fileQueries}
          .showMax=${this.showMax}
          .driveId=${this.driveId}
          .siteId=${this.siteId}
          .groupId=${this.groupId}
          .itemId=${this.itemId}
          .itemPath=${this.itemPath}
          render-on-scroll
          @fileSelected="${() => this.onFileSelected()}"
        ></mgt-file-list>
      </div>
    `;
  }

  /**
   * Render the button used to invoke the flyout
   *
   * @protected
   * @returns {TemplateResult}
   * @memberof MgtFilePicker
   */
  protected renderButton(): TemplateResult {
    const buttonText = 'Select a file';

    return html`
      <!-- <fluent-design-system-provider use-defaults>
        <fluent-button appearance="neutral">Button</fluent-button>
      </fluent-design-system-provider> -->
      <div class="button" @click=${() => this.toggleFlyout()}>
        <div class="button__icon">&#128206;</div>
        <div class="button__text">${buttonText}</div>
      </div>
    `;
  }

  /**
   * Toggle the flyout visiblity
   *
   * @protected
   * @memberof MgtFilePicker
   */
  protected toggleFlyout(): void {
    if (!Providers.globalProvider || Providers.globalProvider.state === ProviderState.SignedOut) {
      return;
    }

    if (this.flyout.isOpen) {
      this.flyout.close();
    } else {
      // Lazy load
      if (!this._doLoad) {
        this._doLoad = true;
        this.requestStateUpdate();
      }

      this.flyout.open();
    }
  }

  /**
   * Handle the event when the user clicks to see all items.
   *
   * @protected
   * @param {PointerEvent} e
   * @memberof MgtFilePicker
   */
  protected handleSeeAll(e: PointerEvent): void {
    this.openFullPicker();
  }

  /**
   * Open the full OneDrive File Picker
   *
   * @protected
   * @memberof MgtFilePicker
   */
  protected openFullPicker(): void {
    console.log('Full picker.');
  }

  /**
   * load state into the component.
   *
   * @protected
   * @returns
   * @memberof MgtFilePicker
   */
  protected async loadState() {
    const provider = Providers.globalProvider;
    if (!provider || provider.state === ProviderState.Loading) {
      return;
    }

    if (provider.state === ProviderState.SignedOut) {
      this.files = null;
      return;
    }
    const graph = provider.graph.forComponent(this);
    let files;

    const getFromMyDrive = !this.driveId && !this.siteId && !this.groupId;

    if (
      (this.driveId && (!this.itemId && !this.itemPath)) ||
      (this.groupId && (!this.itemId && !this.itemPath)) ||
      (this.siteId && (!this.itemId && !this.itemPath))
    ) {
      this.files = null;
    }

    if (!this.files) {
      if (this.fileListQuery) {
        files = await getFilesByListQuery(graph, this.fileListQuery);
      } else if (this.fileQueries) {
        files = await getFilesByQueries(graph, this.fileQueries);
      } else if (getFromMyDrive) {
        if (this.itemId) {
          files = await getFilesById(graph, this.itemId);
        } else if (this.itemPath) {
          files = await getFilesByPath(graph, this.itemPath);
        } else if (this.insightType) {
          files = await getMyInsightsFiles(graph, this.insightType);
        } else {
          files = await getFiles(graph);
        }
      } else if (this.driveId) {
        if (this.itemId) {
          files = await getDriveFilesById(graph, this.driveId, this.itemId);
        } else if (this.itemPath) {
          files = await getDriveFilesByPath(graph, this.driveId, this.itemPath);
        }
      } else if (this.groupId) {
        if (this.itemId) {
          files = await getGroupFilesById(graph, this.groupId, this.itemId);
        } else if (this.itemPath) {
          files = await getGroupFilesByPath(graph, this.groupId, this.itemPath);
        }
      } else if (this.siteId) {
        if (this.itemId) {
          files = await getSiteFilesById(graph, this.siteId, this.itemId);
        } else if (this.itemPath) {
          files = await getSiteFilesByPath(graph, this.siteId, this.itemPath);
        }
      }

      this.files = files;
    }

    // Reset the selected item if it doesn't match any of the new results.
    if (this._selectedItem && this.files.findIndex(v => v.id === this._selectedItem.id) === -1) {
      this._selectedItem = null;
    }
  }

  /**
   * Handle file selection event
   *
   */
  private onFileSelected() {
    const selectedFile = (this.renderRoot.querySelector('#file-list') as MgtFileList).selectedItem;
    this._selectedItem = selectedFile;
    this.fireCustomEvent('fileSelected', this.selectedItem);
  }
}
