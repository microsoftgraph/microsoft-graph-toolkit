import { css } from 'lit-element';

export const icons = css`
  /* @font-face {
    font-family: 'FabricMDL2Icons';
    src: url('../fonts/fabric-icons.woff') format('woff');
  } */

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
