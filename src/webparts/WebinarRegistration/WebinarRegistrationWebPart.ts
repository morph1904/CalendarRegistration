import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'WebinarRegistrationWebPartStrings';
import WebinarRegistration from './components/WebinarRegistration';
import { IWebinarRegistrationProps } from './components/IWebinarRegistrationProps';
import { IWebinarRegistrationState } from './components/IWebinarRegistrationState';

import {setup as pnpSetup } from "@pnp/common";
import { IFilePickerResult, PropertyFieldFilePicker } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/sites";
import { IContextInfo } from "@pnp/sp/sites";
import { PageContext } from '@microsoft/sp-page-context';

var CurrentURL ="";

export interface IWebinarRegistrationWebPartProps {
  flowURL: string;
  btnText: string;
  backgroundImage: IFilePickerResult;
  backgroundImageUrl: string;
  pageContext: PageContext;
}
export default class WebinarRegistrationWebPart extends BaseClientSideWebPart<IWebinarRegistrationWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IWebinarRegistrationProps> = React.createElement(
      WebinarRegistration,
      {
        flowURL: this.properties.flowURL,
        btnText: this.properties.btnText,
        backgroundImage: this.properties.backgroundImage,
        backgroundImageUrl: this.properties.backgroundImageUrl,
        http: this.context.httpClient,
        pageContext: this.context.pageContext,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
      });

    });

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async saveIntoSharePoint(file: IFilePickerResult){
    if (file.fileAbsoluteUrl == null){
      file.downloadFileContent().then(async r => {
        let fileresult = await sp.web.getFolderByServerRelativeUrl(CurrentURL + '/SiteAssets/').files.add(file.fileName, r, true);
        let image = document.location.origin + fileresult.data.ServerRelativeUrl;
        console.log(image);
        this.properties.backgroundImageUrl = image;
      });
  }
  else {
    let image = file.fileAbsoluteUrl;
    console.log(image);
  }
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
                PropertyPaneTextField('flowURL', {
                  label: strings.FlowURLFieldLabel
                }),
                PropertyPaneTextField('btnText', {
                  label: strings.BtnTextFieldLabel
                }),
                PropertyFieldFilePicker('backgroundImage', {
                  context: this.context,
                  filePickerResult: this.properties.backgroundImage,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { this.saveIntoSharePoint(e); this.properties.backgroundImage = e; this.properties.backgroundImageUrl = this.properties.backgroundImage.fileAbsoluteUrl; },
                  onChanged: (e: IFilePickerResult) => { this.saveIntoSharePoint(e); this.properties.backgroundImage = e; this.properties.backgroundImageUrl = this.properties.backgroundImage.fileAbsoluteUrl; },
                  key: "filePickerId",
                  buttonLabel: "File Picker",
                  label: "File Picker",
              })
              ]
            }
          ]
        }
      ]
    };
  }
}
