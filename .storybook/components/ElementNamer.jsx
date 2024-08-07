import { use } from '@esm-bundle/chai';
import React, { useEffect } from 'react';

export const TableNamer = ({ names }) => {
  useEffect(() => {
    const tables = document.querySelectorAll('table:not([aria-hidden="true"])');
    if (tables.length !== names.length) {
      console.error(
        '🦒: TableNamer: number of tables does not match number of names',
        `Found ${tables?.length ?? 0} and was provided ${names?.length ?? 0}`
      );
    }
    for (let i = 0; i < tables.length; i++) {
      tables[i].setAttribute('title', names[i]);
    }
  }, [names]);
  return <></>;
};

export const CopyButtonNamer = ({ names }) => {
  const windowScrollHandler = () => {
    const buttons = document.getElementsByClassName('css-1fdphfk');
    if (buttons) {
      if (buttons.length !== names.length) {
        console.error(
          '🦒: CopyButtonNamer: number of buttons does not match number of names',
          `Found ${buttons?.length ?? 0} buttons and was provided ${names?.length ?? 0} names`
        );
        return;
      }

      for (let i = 0; i < buttons.length; i++) {
        buttons[i].setAttribute('aria-label', names[i]);
      }
      // avoid triggering the scroll event again if you keep scrolling.
      window.removeEventListener('scroll', windowScrollHandler);
    }
  };

  useEffect(() => {
    // NOTE: a good event here is load or DOMContentLoaded BUT for unknown
    // reasons, they don't trigger on the page. The scroll event is a hack
    // based on the fact that to reach the area the code needs an update,
    // you have to scroll - using a mouse or tabbing.
    window.addEventListener('scroll', windowScrollHandler);
    return () => window.removeEventListener('scroll', windowScrollHandler);
  }, [names]);
  return <></>;
};
