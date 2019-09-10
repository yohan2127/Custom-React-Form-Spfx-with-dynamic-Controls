import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { mynewnumber, MYchoices} from '../../models';
import * as strings from 'ProjectTrackingWebPartStrings';
import ProjectTracking from './components/ProjectTracking';
import { IProjectTrackingProps } from './components/IProjectTrackingProps';
import "@pnp/polyfill-ie11";
import { sp, Web } from '@pnp/sp'; 
export interface IProjectTrackingWebPartProps {
  description: string;
  
}

export default class ProjectTrackingWebPart extends BaseClientSideWebPart<IProjectTrackingWebPartProps> {
 // private _newnumber: mynewnumber[]=[];
 // private _opchoices: MYchoices[]=[];
 ///private _opchoices2: MYchoices[]=[];
  ///private _opchoices3: MYchoices[]=[];
 /// private _opchoices4: MYchoices[]=[];

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error?: any) => void): void => {
      sp.setup({
        sp: {
          headers: {
            "Accept": "application/json; odata=nometadata"
          }
        }
      });
      resolve();
    });
  }
  public render(): void {
    if (!this.renderedOnce) {
      //this._drpdown();
      
    }
    const element: React.ReactElement<IProjectTrackingProps > = React.createElement(
      ProjectTracking,
      {
        description: this.properties.description,
        context: this.context 
       // mychoices:this._opchoices,
       // mychoices2:this._opchoices2,
      //  mychoices3:this._opchoices3,
       // mychoices4:this._opchoices4
      }
    );

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
