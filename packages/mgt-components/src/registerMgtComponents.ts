/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import {
  registerMgtAgendaComponent,
  registerMgtFileComponent,
  registerMgtFileListComponent,
  registerMgtGetComponent,
  registerMgtLoginComponent,
  registerMgtPeopleComponent,
  registerMgtPeoplePickerComponent,
  registerMgtPersonCardComponent,
  registerMgtPersonComponent,
  registerMgtPickerComponent,
  registerMgtSearchBoxComponent,
  registerMgtSearchResultsComponent,
  registerMgtSpinnerComponent,
  registerMgtPlannerComponent,
  registerMgtTaxonomyPickerComponent,
  registerMgtTeamsChannelPickerComponent,
  registerMgtThemeToggleComponent,
  registerMgtTodoComponent
} from './components/components';

export const registerMgtComponents = () => {
  // this should match the set of components listed for export in packages/mgt-react/scripts/generate.js
  // all "internal" components should be registered from their parent components
  registerMgtPersonComponent();
  registerMgtPersonCardComponent();
  registerMgtAgendaComponent();
  registerMgtGetComponent();
  registerMgtLoginComponent();
  registerMgtPeoplePickerComponent();
  registerMgtPeopleComponent();
  registerMgtPlannerComponent();
  registerMgtTeamsChannelPickerComponent();
  registerMgtTodoComponent();
  registerMgtFileComponent();
  registerMgtFileListComponent();
  registerMgtPickerComponent();
  registerMgtTaxonomyPickerComponent();
  registerMgtThemeToggleComponent();
  registerMgtSearchBoxComponent();
  registerMgtSearchResultsComponent();
  registerMgtSpinnerComponent();
};
