/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { html } from 'lit';
import { withCodeEditor } from '../../../.storybook/addons/codeEditorAddon/codeAddon';

export default {
  title: 'Components / mgt-taxonomy-picker / React',
  component: 'taxonomy-picker',
  decorators: [withCodeEditor]
};

export const TaxonomyPicker = () => html`
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
  <react>
    import { TaxonomyPicker } from '@microsoft/mgt-react';

    export default () => (
      <TaxonomyPicker termSetId="f1c3d275-b202-41f0-83f3-80d63ffaa052"></TaxonomyPicker>
    );
  </react>
`;

export const Events = () => html`
  <!-- Check the console tab for results -->
  <mgt-taxonomy-picker term-set-id="f1c3d275-b202-41f0-83f3-80d63ffaa052"></mgt-taxonomy-picker>
  <react>
    // Check the console tab for the event to fire
    import { useCallback } from 'react';
    import { TaxonomyPicker } from '@microsoft/mgt-react';

    export default () => {
      const onUpdated = useCallback((e) => {
        console.log('updated', e); 
      });
      const onSelectionChanged = useCallback((e: CustomEvent<IDynamicPerson[]>) => {
        console.log(e.detail);
      }, []);

      return (
        <TaxonomyPicker 
        termSetId="f1c3d275-b202-41f0-83f3-80d63ffaa052" 
        updated={onUpdated}
        selectionChanged={onSelectionChanged}>
      </TaxonomyPicker>
      );
    };
  </react>
  <script>
    document.querySelector('mgt-taxonomy-picker').addEventListener('selectionChanged', e => {
      console.log('selected term:', e.detail)
    });
  </script>
`;
