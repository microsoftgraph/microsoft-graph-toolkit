import { BetaGraph, IGraph, Providers } from '@microsoft/mgt-element';
import { produce } from 'immer';
import { StatefulClient } from './StatefulClient';
import { graph } from '../utils/graph';
import { GraphConfig } from './GraphConfig';

export abstract class BaseStatefulClient<T> implements StatefulClient<T> {
  private _graph: IGraph | undefined;

  protected get graph(): IGraph | undefined {
    if (this._graph) return this._graph;
    if (Providers.globalProvider?.graph) {
      this._graph = graph('mgt-chat', GraphConfig.version);
    }
    return this._graph;
  }

  protected set graph(value: IGraph | undefined) {
    this._graph = value;
  }

  protected get betaGraph(): BetaGraph | undefined {
    const g = this.graph;
    if (g) return BetaGraph.fromGraph(g);

    return undefined;
  }

  private _subscribers: ((state: T) => void)[] = [];
  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public onStateChange(handler: (state: T) => void): void {
    if (!this._subscribers.includes(handler)) {
      this._subscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving state updates
   *
   * @param {(state: GraphChatClient) => void} handler
   * @memberof StatefulGraphChatClient
   */
  public offStateChange(handler: (state: T) => void): void {
    const index = this._subscribers.indexOf(handler);
    if (index !== -1) {
      this._subscribers = this._subscribers.splice(index, 1);
    }
  }

  /**
   * Calls each subscriber with the next state to be emitted
   *
   * @param recipe - a function which produces the next state to be emitted
   */
  protected notifyStateChange(recipe: (draft: T) => void) {
    this.state = produce(this.state, recipe);
    this._subscribers.forEach(handler => handler(this.state));
  }

  protected abstract state: T;
  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatClient}
   * @memberof StatefulGraphChatClient
   */
  public getState(): T {
    return this.state;
  }
}
