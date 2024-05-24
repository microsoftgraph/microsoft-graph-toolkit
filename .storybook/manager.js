/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { addons, types } from '@storybook/manager-api';
import theme from './theme';

import React, { useState } from 'react';
import { useChannel } from '@storybook/api';
import { Providers } from '../packages/mgt-element/dist/es6/providers/Providers';
import { ProviderState, LoginType } from '../packages/mgt-element/dist/es6/providers/IProvider';
import { Msal2Provider } from '../packages/providers/mgt-msal2-provider/dist/es6/Msal2Provider';
import { CLIENTID, SETPROVIDER_EVENT, AUTH_PAGE } from './env';
import { MockProvider, PACKAGE_VERSION } from '@microsoft/mgt-element';
import { Login, Person } from '../packages/mgt-react/src/generated/react';
import './manager.css';

const getClientId = () => {
  const urlParams = new window.URL(window.location.href).searchParams;
  const customClientId = urlParams.get('clientId');

  return customClientId ? customClientId : CLIENTID;
};

document.getElementById('mgt-version').innerText = PACKAGE_VERSION;

const mockProvider = new MockProvider(true);
const msal2Provider = new Msal2Provider({
  clientId: getClientId(),
  redirectUri: window.location.origin + '/' + AUTH_PAGE,
  scopes: [
    // capitalize all words in the scope
    'user.read',
    'user.read.all',
    'mail.readBasic',
    'people.read',
    'people.read.all',
    'sites.read.all',
    'user.readbasic.all',
    'contacts.read',
    'presence.read',
    'presence.read.all',
    'tasks.readwrite',
    'tasks.read',
    'calendars.read',
    'group.read.all',
    'files.read',
    'files.read.all',
    'files.readwrite',
    'files.readwrite.all'
  ],
  loginType: LoginType.Popup
});
let loginInitiated = false;

Providers.globalProvider = msal2Provider;

const SignInPanel = () => {
  const [state, setState] = useState(Providers.globalProvider.state);

  const emit = useChannel({
    STORY_RENDERED: id => {
      console.log('storyRendered', id);
    }
  });

  const emitProvider = loginState => {
    if (Providers.globalProvider.state === ProviderState.SignedOut && Providers.globalProvider !== mockProvider) {
      emit(SETPROVIDER_EVENT, { state: loginState, provider: mockProvider, name: 'MgtMockProvider' });
    } else {
      emit(SETPROVIDER_EVENT, { state: loginState, provider: msal2Provider, name: 'MgtMsal2Provider' });
    }
  };

  Providers.onProviderUpdated(() => {
    if (state === Providers.globalProvider.state) return;
    setState(Providers.globalProvider.state);
    emitProvider(Providers.globalProvider.state);
  });

  const onSignOut = async () => {
    await Providers.globalProvider.logout();
    reload();
  };

  const reload = () => {
    console.log('reload');
    window.location.reload();
  };

  const onLoginCompleted = e => {
    if (loginInitiated) {
      reload();
    }
    console.log('loginCompleted');
  };

  const onLoginInitiated = e => {
    loginInitiated = true;
    console.log('loginInitiated');
  };

  emitProvider(state);

  return (
    <>
      <div
        style={{
          marginTop: '3px',
          display: Providers.globalProvider.state !== ProviderState.SignedIn ? 'flex' : 'none'
        }}
      >
        {/* We need to keep the login component available (but hidden) to handle mock to logged in states */}
        <Login loginView="compact" loginCompleted={onLoginCompleted} loginInitiated={onLoginInitiated}></Login>
      </div>

      <div style={{ display: Providers.globalProvider.state === ProviderState.SignedIn ? 'flex' : 'none' }}>
        <div style={{ marginTop: '8px' }}>
          <Person personQuery="me"></Person>
        </div>
        <fluent-button appearance="lightweight" style={{ marginTop: '3px' }} onClick={onSignOut}>
          Sign Out
        </fluent-button>
      </div>
    </>
  );
};

addons.register('mgt', api => {
  addons.add('mgt/login', {
    type: types.TOOLEXTRA,
    title: 'Sign In',
    render: ({ active }) => <SignInPanel />
  });
});

addons.setConfig({
  enableShortcuts: false,
  theme,
  sidebar: {
    filters: {
      patterns: item => {
        return !(item.tags.includes('hidden') && item.type === 'story');
      }
    }
  }
});

// inject page setup for manager frame.

const xmlns = 'http://www.w3.org/2000/svg';

window.onload = () => {
  addUsefulLinks();
};
// Dirty hack. The screen sizing btn resets the title on clicking it. When you
// set the correct title during the window loading event, after clicking the
// button for a full screen or exit of a full screen, the set title is 
// overridden to the previous problematic value. This is a catch-all hack on the
// window click event to just ensure any clicks on the btn resets to the correct title.
window.onclick = () =>{
  resetScreenSizingBtnTitles()
}

function addUsefulLinks() {
  const linkStyle = 'color: black; text-decoration: none; text-align: center; cursor: pointer; padding-right: 0.5rem;';
  const textStyle = 'margin-left: 0.25rem; font-size: 0.75rem;';
  const linkContentStyle = 'display: flex;';

  const sidebarHeader = document.getElementsByClassName('sidebar-header');
  if (sidebarHeader.length === 0) {
    // sidebar container has not loaded in the page yet, retry in 500ms
    setTimeout(addUsefulLinks, 500);
    return false;
  }

  // This is a fix for an accessibility issue: https://github.com/microsoftgraph/microsoft-graph-toolkit/issues/1085
  sidebarHeader[0].innerHTML =
    '<h1 tabindex="-1" class="css-1su1ft1">' +
    '<a href="https://aka.ms/mgt" target="_blank" class="css-ixbm00">' +
    'Microsoft Graph Toolkit Playground</a>' +
    '</h1>';

  const sidebarNode = sidebarHeader[0].parentNode;

  // Github
  const ghSvgPath =
    'M12 .297c-6.63 0-12 5.373-12 12 0 5.303 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 22.092 24 17.592 24 12.297c0-6.627-5.373-12-12-12';
  const githubSvgElem = createSvg(ghSvgPath);

  const githubText = document.createElement('span');
  githubText.style = textStyle;
  githubText.innerText = 'GitHub';

  const repoLinkContent = document.createElement('span');
  repoLinkContent.style = linkContentStyle;
  repoLinkContent.append(githubSvgElem, githubText);

  const repoLink = document.createElement('a');
  repoLink.href = 'https://aka.ms/mgt';
  repoLink.target = '_blank';
  repoLink.style = linkStyle;
  repoLink.appendChild(repoLinkContent);

  // npm
  const npmSvgPath =
    'M1.763 0C.786 0 0 .786 0 1.763v20.474C0 23.214.786 24 1.763 24h20.474c.977 0 1.763-.786 1.763-1.763V1.763C24 .786 23.214 0 22.237 0zM5.13 5.323l13.837.019-.009 13.836h-3.464l.01-10.382h-3.456L12.04 19.17H5.113z';
  const npmSvgElem = createSvg(npmSvgPath);

  const npmText = document.createElement('span');
  npmText.style = textStyle;
  npmText.innerText = 'npm';

  const npmLinkContent = document.createElement('span');
  npmLinkContent.style = linkContentStyle;
  npmLinkContent.append(npmSvgElem, npmText);

  const npmLink = document.createElement('a');
  npmLink.href = 'https://www.npmjs.com/package/@microsoft/mgt';
  npmLink.target = '_blank';
  npmLink.style = linkStyle + ' margin-left: 0.5rem;';
  npmLink.appendChild(npmLinkContent);

  // links container
  const usefulLinksContainer = document.createElement('div');
  usefulLinksContainer.id = 'useful-links';
  usefulLinksContainer.style = 'display: flex; margin: 1rem 0rem; font-size: 0.875rem;';

  usefulLinksContainer.append(repoLink, npmLink);

  sidebarNode.insertBefore(usefulLinksContainer, sidebarNode.childNodes[1]);

  const sidebarSubheading = document.getElementsByClassName('sidebar-subheading');
  if (sidebarSubheading) {
    for (let i = 0; i < sidebarSubheading.length; i++) {
      const subheading = sidebarSubheading[i];
      subheading.removeAttribute('aria-expanded');
    }
  }

  function setEventOnMenuClick() {
    const expandCollapseMenu = document.getElementsByClassName('css-ulso1l');
    if (expandCollapseMenu) {
      for (let i = 0; i < expandCollapseMenu.length; i++) {
        const menu = expandCollapseMenu[i];
        menu.addEventListener('click', setArialLabelForExpandCollapseBtn, { useCapture: true });
      }
    }
  }

  function setArialLabelForExpandCollapseBtn() {
    const expandCollapseBtns = document.getElementsByClassName('css-rl1ij0');
    if (expandCollapseBtns) {
      for (let i = 0; i < expandCollapseBtns.length; i++) {
        const button = expandCollapseBtns[i];
        setButtonAriaLabel(button);
        button.addEventListener('click', btnUpDown, { useCapture: true });
      }
    }
  }

  function btnUpDown(event) {
    const element = event.target; // button when using tabs to navigate, svg when using mouse.
    setButtonAriaLabel(element);
  }

  function setButtonAriaLabel(element) {
    const dataExpandedState = element.getAttribute('data-expanded');
    const ariaValue = dataExpandedState === 'true' ? 'expand' : 'collapse';
    element.setAttribute('aria-label', ariaValue);
  }

  setEventOnMenuClick();
  resetScreenSizingBtnTitles();
}

function createSvg(svgPath) {
  const boxWidth = 12;
  const boxHeight = 12;
  const imageStyle = 'height: 0.75rem; margin-top: 0.125rem;';
  const svgElem = document.createElementNS(xmlns, 'svg');
  svgElem.setAttributeNS(null, 'viewBox', `0 0 ${boxWidth * 2} ${boxHeight * 2}`);
  svgElem.setAttributeNS(null, 'width', boxWidth);
  svgElem.setAttributeNS(null, 'height', boxHeight);
  svgElem.setAttributeNS(null, 'role', 'img');
  svgElem.setAttributeNS(null, 'aria-label', ' to ');
  svgElem.setAttributeNS(null, 'style', imageStyle);

  const svgElemPath = document.createElementNS(xmlns, 'path');
  svgElemPath.setAttributeNS(null, 'd', svgPath);

  svgElem.appendChild(svgElemPath);
  return svgElem;
}

function resetScreenSizingBtnTitles () {
  const titles = {
    'Exit full screen [F]': 'Exit full screen',
    'Go full screen [F]': 'Go full screen'
  }

  const screenSizingBtns = document.getElementsByClassName('css-6jc9zw');
  for (let i = 0; i < screenSizingBtns.length; i++) {
    const btn = screenSizingBtns[i];
    const title = btn.getAttribute('title');
    // zooming in/out btns share the same class, prefer their titles when they
    // are not found in the titles object above.
    btn.setAttribute('title', titles[title] ?? title);
    btn.setAttribute('aria-label', titles[title] ?? title);
  }
}

const injectScripts = () => {
  const consentScript = document.createElement('script');
  consentScript.src = 'https://consentdeliveryfd.azurefd.net/mscc/lib/v2/wcp-consent.js';
  document.head.appendChild(consentScript);
  const fluentUI = document.createElement('script');
  fluentUI.type = 'module';
  fluentUI.src = 'https://unpkg.com/@fluentui/web-components@2';
  document.head.appendChild(fluentUI);
};

injectScripts();

let attempts = 0;
// Uses setInterval to wait for the sidebar to be rendered into the page
// interval is cleared on success or after 100 tries (10 seconds)
const retry = setInterval(() => {
  attempts++;
  if (attempts > 100) {
    clearInterval(retry);
    return;
  }
  const footer = document.querySelector('footer');
  const sidebar = document.querySelector('nav.sidebar-container');
  if (!sidebar) return;
  sidebar.parentElement.appendChild(footer);
  sidebar.style.height = `calc(100% - ${footer.clientHeight}px)`;
  clearInterval(retry);
}, 100);

navigator.serviceWorker.register('sw.js');
