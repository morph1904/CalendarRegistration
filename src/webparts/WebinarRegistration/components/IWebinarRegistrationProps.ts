import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';
import { HttpClient } from "@microsoft/sp-http";
import { PageContext } from '@microsoft/sp-page-context';
export interface IWebinarRegistrationProps {
  flowURL: string;
  btnText: string;
  backgroundImage: IFilePickerResult;
  backgroundImageUrl: string;
  http:HttpClient;
  pageContext: PageContext;

}
