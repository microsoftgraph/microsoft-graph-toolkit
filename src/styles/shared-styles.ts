import { css } from 'lit-element';
import { icons } from './fabric-icons';

export const sharedStyles = [
  icons,
  css`
    :host([hidden]) {
      display: none;
    }
    :host {
      display: block;
      --default-font-family: 'Segoe UI', 'Segoe UI Web (West European)',
        'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue',
        sans-serif;
      --default-hover-color: #0078d7;
    }
  `
];
