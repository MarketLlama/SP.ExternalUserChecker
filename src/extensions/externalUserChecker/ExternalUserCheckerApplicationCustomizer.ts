import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, 
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { sp } from "@pnp/sp";
import { Dialog } from '@microsoft/sp-dialog';
import styles from './AppCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { PermissionKind } from "@pnp/sp";
import * as strings from 'ExternalUserCheckerApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ExternalUserCheckerApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IExternalUserCheckerApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
  Bottom: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ExternalUserCheckerApplicationCustomizer
  extends BaseApplicationCustomizer<IExternalUserCheckerApplicationCustomizerProperties> {

   // These have been added
   private _topPlaceholder: PlaceholderContent | undefined;
   private _rendered : boolean = false;
  @override
  public onInit(): Promise<void> { 
    
    return super.onInit().then(_ => {

      sp.setup({
        spfxContext: this.context
      });

      Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

      // Added to handle possible changes on the existence of placeholders.
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
  
      // Call render method for generating the HTML elements.
      this._renderPlaceHolders();
      
    });
      //return Promise.resolve<void>();
  }
  private _renderPlaceHolders = () => {
    console.log(
        "Available placeholders: ",
        this.context.placeholderProvider.placeholderNames
            .map(name => PlaceholderName[name])
            .join(", ")
    );

    // Handling the top placeholder
    if (!this._topPlaceholder) {
        this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
            PlaceholderName.Top,
            { onDispose: this._onDispose }
        );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
          console.error("The expected placeholder (Top) was not found.");
          return;
      }
      //set a internal checking for external users.
      if(!this._rendered){
        setInterval(this._checkUserIsExternal,5000);
      }
    }
  }
  //Checks thst the user is not a part of the company.
  private _checkUserIsExternal = async() =>{  
    if(this._rendered){
      return;
    } 
    let currentUser = await sp.web.currentUser.get();
      //if user is not external then check to see if other users are...
    if(!currentUser.IsShareByEmailGuestUser || !currentUser.IsEmailAuthenticationGuestUser) {
      let siteUsers = await sp.web.siteUsers.get();
      siteUsers.forEach(async user=>{
        if (user.IsShareByEmailGuestUser || user.IsEmailAuthenticationGuestUser){
          let hasExternalUsersWithPermissions = await this._checkUserPermisions(user);
          //if site has external users with permission to view the site then show the message
          if(hasExternalUsersWithPermissions){
            this._showMessage();
            this._rendered = true;
          }
          return;
        }
      });
    }
  }
  //Checks user has permission to view pages within a site.
  private _checkUserPermisions = async(user) : Promise<any> => {
    return new Promise(resolve =>{
      sp.web.userHasPermissions(user.LoginName,PermissionKind.ViewPages).then(p =>{
        resolve(p);
      });
    });
  }
  //Shows Ribbon Message within a teamsite =
  private _showMessage = () =>{
    if (this.properties) {
      let topString: string = this.properties.Top;
      if (!topString) {
          topString = "(Top property was not defined.)";
      }
      if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
              <div class="${styles.top}">
                  <i class="${styles.warningIcon} ms-Icon ms-Icon--Warning" aria-hidden="true"></i><p> ${escape(
                      topString
                  )}</P>
              </div>
          </div>`;
      }
    }
  }
  private _onDispose = () => {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }
}
