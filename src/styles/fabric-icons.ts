import './fabric-icon-font';
import { css } from 'lit-element';

export const icons = css`
  .ms-Icon {
    display: inline-block;
    font-family: 'FabricMDL2Icons';
    font-style: normal;
    font-weight: normal;
    -moz-osx-font-smoothing: grayscale;
    -webkit-font-smoothing: antialiased;
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
