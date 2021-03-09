import * as React from 'react';
import styles from './WebinarRegistration.module.scss';
import { IWebinarRegistrationProps } from './IWebinarRegistrationProps';
import { IWebinarRegistrationState } from './IWebinarRegistrationState';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { escape } from '@microsoft/sp-lodash-subset';
import { PropertyFieldFilePicker, IPropertyFieldFilePickerProps, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';
import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { sp } from '@pnp/sp/presets/all';


export default class WebinarRegistration extends React.Component<IWebinarRegistrationProps, IWebinarRegistrationState, {}> {
  public constructor(props: IWebinarRegistrationProps, state: IWebinarRegistrationState) {
    super(props);

    this.state = {
      user: "",
      webinar: {},
      hideDialog: true,
      dialogText: "",
      dialogType: "",
      isLoading: false,
      webinarExpired:false
    };
  }
  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
    this.render();
  }
  public WebinarRegister = () => {
    this.setState({
      isLoading: true
    });
    console.log(this.state.webinar.ID);
    let queryString: string = JSON.stringify({
      'userUPN': this.state.user,
      'WebinarID': this.state.webinar.WebinarID.ID,
      'pageID': this.state.webinar.ID
    });

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Content-type', 'application/json');
    requestHeaders.append('Cache-Control', 'no-cache');
    const httpClientOptions: IHttpClientOptions = {
      body: queryString,
      headers: requestHeaders
    };

    return this.props.http.post(
      this.props.flowURL,
      HttpClient.configurations.v1,
      httpClientOptions
    ).then(
      (response: HttpClientResponse) =>{

        return response.json();
      }).then(jsonResponse => {
        console.log(jsonResponse);
        if(jsonResponse.error){
          this.setState({
            isLoading: false,
            dialogType: "Error!",
            dialogText: jsonResponse.error
          });
          return this._showDialog();
        }
        if (jsonResponse.success){
          this.setState({
            isLoading: false,
            dialogType: "Success!",
            dialogText: jsonResponse.success
          });
          return this._showDialog();
        }

        return jsonResponse;
      });
  }
  public componentDidMount() {
    let CurrentURL = this.props.pageContext.site.serverRelativeUrl;
    sp.web.lists.getById(this.props.pageContext.list.id.toString())
      .items.getById(this.props.pageContext.listItem.id)
        .select("Webinar_x0020_Title","id", "WebinarID/ID","WebinarID/StartDateAndTime")
        .expand("WebinarID")
          .get()
            .then(d => {
              this.setState({
                  webinar: d
              });
              console.log(this.state.webinar);
              var currentdate = new Date();
              if (this.state.webinar.WebinarID.StartDateAndTime <= currentdate.toISOString()){

                this.setState({
                  webinarExpired: true
                });
              }
              console.log(this.state.webinarExpired);
            });
            sp.web.currentUser.get().then(f => {
              this.setState({
                user: f.UserPrincipalName
            });
            });
  }

  public render(): React.ReactElement<IWebinarRegistrationProps> {
    let imageUrl;
    //console.log(this.props.backgroundImageUrl)
    if(this.props.backgroundImageUrl){
       imageUrl = this.props.backgroundImageUrl;
    }else {
       imageUrl = require('../../assets/defaultBackground.jpg');
    }

    var style = {
      backgroundImage: `url(${imageUrl})`,
      backgroundSize: "cover",
    };

    return (
      <div className={ styles.WebinarRegistration }>
        <div className={ styles.container } id="main-container">
        {this.state.isLoading ? (
          <Spinner label="Registering Please wait..." ariaLive="assertive" labelPosition="bottom" className={styles.loadingbar}/>):(
          <div className={ styles.row } style={style}>
            <div className={ styles.column }>
              <h1 className={ styles.title }>Register for {this.state.webinar.Webinar_x0020_Title}</h1>
              {this.state.webinarExpired ? (
              <button onClick={this.WebinarRegister} className={ styles.button } disabled={true}>
              <span className={ styles.label }>Webinar has Expired</span>
            </button>):(
              <button onClick={this.WebinarRegister} className={ styles.button }>
              <span className={ styles.label }>{this.props.btnText}</span>
            </button>
            )}
            </div>
          </div>
        )}
        </div>
        <Dialog
          hidden={this.state.hideDialog}
          onDismiss={this._closeDialog}
          dialogContentProps={{
            type: DialogType.close,
            title: this.state.dialogType,
            subText:this.state.dialogText,
          }}

          modalProps={{
            isBlocking: false,
            styles: { main: { maxWidth: 450 } },
          }}
        >
          <DialogFooter>
            <PrimaryButton onClick={this._closeDialog} text="Close" />
          </DialogFooter>
        </Dialog>
      </div>

    );
  }
}
