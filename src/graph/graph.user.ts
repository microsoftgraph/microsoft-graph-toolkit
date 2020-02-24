import { User } from '@microsoft/microsoft-graph-types';
import { IGraph } from '../IGraph';
import { prepScopes } from '../utils/GraphHelpers';

/**
 * async promise, returns Graph User data relating to the user logged in
 *
 * @returns {(Promise<User>)}
 * @memberof Graph
 */
export function getMe(graph: IGraph): Promise<User> {
  return graph
    .api('me')
    .middlewareOptions(prepScopes('user.read'))
    .get();
}
