import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import styles from './Sp2010AlertBannerAppExtentionApplicationCustomizer.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'Sp2010AlertBannerAppExtentionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'Sp2010AlertBannerAppExtentionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISp2010AlertBannerAppExtentionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string;
  Bottom: string;
}


/** A Custom Action which can be run during execution of a Client Side Application */
export default class Sp2010AlertBannerAppExtentionApplicationCustomizer
  extends BaseApplicationCustomizer<ISp2010AlertBannerAppExtentionApplicationCustomizerProperties> {

  // These have been added
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);



    // Wait for the placeholders to be created (or handle them being changed) and then
    // render.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
  
    return Promise.resolve();

    /*
    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    return Promise.resolve();
    */
  }

  private _renderPlaceHolders(): void {
    console.log("Sp2010AlertBannerAppExtentionApplicationCustomizer._renderPlaceHolders()");
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
  
      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.top}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>
              This site contains 2010 Workflow(s) which will cease to function on 12/31/2021. Please, click <a href="https://dvagov.sharepoint.com/sites/OITSharePointPlatform/SitePages/2010-Workflow.aspx?source=/sites/OITSharePointPlatform/_layouts/15/news.aspx&promotedState=1">here</a> to locate the workflows. If you need assistance modernizing the 2010 Workflow(s), please, contact your SharePoint Site Admin or Site Owner.
            </div>
          </div>`;
        }
      }
    }
    let bottomMessage:boolean = false;
    // Handling the bottom placeholder
    if (bottomMessage && !this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }
  
      if (this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          bottomString = "(Bottom property was not defined.)";
        }
  
        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.bottom}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                bottomString
              )}
            </div>
          </div>`;
        }
      }
    }
  }

  private _onDispose(): void {
    console.log('[Sp2010AlertBannerAppExtentionApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }

}
