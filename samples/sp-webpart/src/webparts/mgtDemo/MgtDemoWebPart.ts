import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { lazyLoadComponent } from '@microsoft/mgt-spfx-utils';

import * as strings from 'MgtDemoWebPartStrings';

// import the providers at the top of the page
import { Providers } from '@microsoft/mgt-element/dist/es6/providers/Providers';
import { customElementHelper } from '@microsoft/mgt-element/dist/es6/components/customElementHelper';
import { SharePointProvider } from '@microsoft/mgt-sharepoint-provider/dist/es6/SharePointProvider';
// Async import of component that imports the React Components
const MgtDemo = React.lazy(() => import(/* webpackChunkName: 'mgt-demo-component' */'./components/MgtDemo'));
export interface IMgtDemoWebPartProps {
  description: string;
}
// set the disambiguation before initializing any web part
customElementHelper.withDisambiguation('mgt-demo-client-side-solution');

export default class MgtDemoWebPart extends BaseClientSideWebPart<IMgtDemoWebPartProps> {
  // set the global provider
  protected async onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
  }

  public render(): void {
    const element = lazyLoadComponent(MgtDemo, { description: this.properties.description });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
