import { IGraph, Providers } from '@microsoft/mgt-element';

export const graph = (component: string): IGraph => {
  if (!component) throw new Error('Component name is required');
  return Providers.globalProvider.graph.forComponent(component);
};
