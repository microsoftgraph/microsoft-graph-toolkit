/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { styles } from './theme-css';

const fabricFont = document.createElement('style');
fabricFont.type = 'text/css';
fabricFont.appendChild(
  document.createTextNode(`
@font-face {
    font-family: 'FabricMDL2Icons';
    src: url('https://static2.sharepointonline.com/files/fabric/assets/icons/fabricmdl2icons-2.68.woff2') format('woff2'),
    url(https://static2.sharepointonline.com/files/fabric/assets/icons/fabricmdl2icons-2.68.woff) format("woff"),
    url(https://static2.sharepointonline.com/files/fabric/assets/icons/fabricmdl2icons-2.68.ttf) format("truetype");;
}
`)
);
document.head!.appendChild(fabricFont);

const themeStyle = document.createElement('style');
themeStyle.innerHTML = styles.toString();
document.head!.appendChild(themeStyle);
