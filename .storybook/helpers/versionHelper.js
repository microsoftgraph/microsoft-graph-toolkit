import { PACKAGE_VERSION } from '../../packages/mgt-element/dist/es6/utils/version';

export const getVersion = async (title, files) => {
  if (PACKAGE_VERSION.toLowerCase().includes('pr') || window.location.href.includes('localhost')) {
    return 'next';
  } else {
    return PACKAGE_VERSION;
  }
};
