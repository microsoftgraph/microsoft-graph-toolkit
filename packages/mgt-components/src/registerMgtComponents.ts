import {
  registerMgtAgendaComponent,
  registerMgtContactComponent,
  registerMgtFileComponent,
  registerMgtLoginComponent,
  registerMgtPeopleComponent,
  registerMgtPersonCardComponent,
  registerMgtPersonComponent,
  registerMgtSpinnerComponent
} from './components/components';
import { registerMgtFlyoutComponent } from './components/sub-components/mgt-flyout/mgt-flyout';

export const registerMgtComponents = () => {
  // this should match the set of components listed for export in packages/mgt-react/scripts/generate.js
  // all "internal" components should be registered from their parent components
  registerMgtAgendaComponent();
  registerMgtContactComponent();
  registerMgtFileComponent();
  registerMgtFlyoutComponent();
  registerMgtLoginComponent();
  registerMgtPeopleComponent();
  registerMgtPersonComponent();
  registerMgtPersonCardComponent();
  registerMgtSpinnerComponent();
};
