import React, { useEffect } from "react";

export const TableNamer = ({names}) => {
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
    }, []);
    return <></>
};
