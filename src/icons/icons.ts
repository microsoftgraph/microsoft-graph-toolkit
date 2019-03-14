import { css } from 'lit-element';
import './icon-font';

export const icons = css`
  .ms-Icon {
    -moz-osx-font-smoothing: grayscale;
    -webkit-font-smoothing: antialiased;
    display: inline-block;
    font-family: 'FabricMDL2Icons';
    font-style: normal;
    font-weight: normal;
  }

  .ms-Icon--ChevronDown::before {
    content: '\\\E70D';
  }

  .ms-Icon--ChevronUp::before {
    content: '\\\E70E';
  }

  .ms-Icon--Contact::before {
    content: '\\\E77B';
  }

  .ms-Icon--AddFriend::before {
    content: '\\\E8FA';
  }

  .ms-Icon--OutlookLogoInverse::before {
    content: '\\\EB6D';
  }
`;
