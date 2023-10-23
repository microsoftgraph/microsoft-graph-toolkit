import { customElementHelper } from '../components/customElementHelper';

export const buildComponentName = (tagBase: string) => `${customElementHelper.prefix}-${tagBase}`;

export const registerComponent = (
  tagBase: string,
  constructor: CustomElementConstructor,
  options?: ElementDefinitionOptions
) => {
  const tagName = buildComponentName(tagBase);
  if (!customElements.get(tagName)) customElements.define(tagName, constructor, options);
};
