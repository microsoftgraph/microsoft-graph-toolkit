import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

/**
 * IDynamicPerson describes the person object we use throughout mgt-person,
 * which can be one of three similar Graph types.
 *
 * In addition, this custom type also defines the optional `personImage` property,
 * which is used to pass the image around to other components as part of the person object.
 */
export type IDynamicPerson = (
  | MicrosoftGraph.User
  | MicrosoftGraph.Person
  | MicrosoftGraph.Contact
  | MicrosoftGraph.Group) & {
  /**
   * personDetails.personImage is a toolkit injected property to pass image between components
   * an optimization to avoid fetching the image when unnecessary.
   *
   * @type {string}
   */
  personImage?: string;
};
