import { provideFluentDesignSystem } from '@fluentui/web-components';

const designSystem = provideFluentDesignSystem();

export const registerFluentComponents = (...fluentComponents) => {
  if (!fluentComponents || !fluentComponents.length) {
    return;
  }

  for (const component of fluentComponents) {
    designSystem.register(component());
  }
};
