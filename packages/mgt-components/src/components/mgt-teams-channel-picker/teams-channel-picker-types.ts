/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/**
 * Team with displayName
 *
 * @export
 * @interface SelectedChannel
 */

export type Team = MicrosoftGraph.Team & {
  /**
   * Display name Of Team
   *
   * @type {string}
   */
  displayName?: string;
};
/**
 * Selected Channel item
 *
 * @export
 * @interface SelectedChannel
 */

export interface SelectedChannel {
  /**
   * Channel
   *
   * @type {MicrosoftGraph.Channel}
   * @memberof SelectedChannel
   */
  channel: MicrosoftGraph.Channel;

  /**
   * Team
   *
   * @type {MicrosoftGraph.Team}
   * @memberof SelectedChannel
   */
  team: Team;
}
/**
 * Drop down menu item
 *
 * @export
 * @interface DropdownItem
 */
export interface DropdownItem {
  /**
   * Teams channel
   *
   * @type {DropdownItem[]}
   * @memberof DropdownItem
   */
  channels?: DropdownItem[];
  /**
   * Microsoft Graph Channel or Team
   *
   * @type {(MicrosoftGraph.Channel | MicrosoftGraph.Team)}
   * @memberof DropdownItem
   */
  item: MicrosoftGraph.Channel | Team;
}
/**
 * Drop down menu item state
 *
 * @interface DropdownItemState
 */
export interface ChannelPickerItemState {
  /**
   * Microsoft Graph Channel or Team
   *
   * @type {(MicrosoftGraph.Channel | MicrosoftGraph.Team)}
   * @memberof ChannelPickerItemState
   */
  item: MicrosoftGraph.Channel | Team;
  /**
   * if dropdown item shows expanded state
   *
   * @type {boolean}
   * @memberof DropdownItemState
   */
  isExpanded?: boolean;
  /**
   * If item contains channels
   *
   * @type {ChannelPickerItemState[]}
   * @memberof DropdownItemState
   */
  channels?: ChannelPickerItemState[];
  /**
   * if Item has parent item (team)
   *
   * @type {ChannelPickerItemState}
   * @memberof DropdownItemState
   */
  parent: ChannelPickerItemState;
}
