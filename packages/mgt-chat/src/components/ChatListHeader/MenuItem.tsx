import type { Slot } from '@fluentui/react-utilities';

export interface MenuItem {
  displayText: string;
  icon?: Slot<'span'>;
  onSelected: React.ReactEventHandler<HTMLAnchorElement>;
}
