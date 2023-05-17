import { PageOpenBehavior } from '../helpers/UrlHelper';
import { ILocalizedString } from './ILocalizedString';

export interface IDataVerticalConfiguration {
  /**
   * Unique key for the vertical
   */
  key: string;

  /**
   * The vertical tab name
   */
  tabName: string | ILocalizedString;

  /**
   * The vertical tab value that will be sent to connected components
   */
  tabValue: string;

  /**
   * Specifes if the vertical is a link
   */
  isLink: boolean;

  /**
   * The link URL
   */
  linkUrl: string;

  /**
   * The link open behavior
   */
  openBehavior: PageOpenBehavior;
}
