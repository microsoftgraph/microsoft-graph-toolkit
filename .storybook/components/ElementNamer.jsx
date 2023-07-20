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
  useEffect(() => {
    const onWindowLoadHander = () => {
      const buttons = document.getElementsByClassName('css-3ltsna');
      if (buttons.length !== names.length) {
        console.error(
          '🦒: CopyButtonNamer: number of buttons does not match number of names',
          `Found ${buttons?.length ?? 0} buttons and was provided ${names?.length ?? 0} names`
        );
      }
      for (let i = 0; i < buttons.length; i++) {
        buttons[i].setAttribute('aria-label', names[i]);
      }
    };
    window.addEventListener('load', onWindowLoadHander);
    return () => window.removeEventListener('load', onWindowLoadHander);
  }, [names]);
  return <></>;
};
