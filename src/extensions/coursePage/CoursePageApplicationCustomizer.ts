import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
  IPlaceholderCreateContentOptions
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CoursePageApplicationCustomizerStrings';

const LOG_SOURCE: string = 'CoursePageApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICoursePageApplicationCustomizerProperties {
  // This is an example; replace with your own property
  title: string;
  url: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class CoursePageApplicationCustomizer
  extends BaseApplicationCustomizer<ICoursePageApplicationCustomizerProperties> {

  private topPlaceHolder: PlaceholderContent;
  private bottomPlaceHolder: PlaceholderContent;


  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.context.placeholderProvider.changedEvent.add(this, this.renderPlaceHolders);

    return Promise.resolve();
  }

  private renderPlaceHolders(): void {

    console.log("Available Placeholders :");
    this.context.placeholderProvider.placeholderNames.forEach((name) => {
      console.log(name.toString());
    });

    if (!this.topPlaceHolder) {

      // Try to Create the Top Place Holder
      this.topPlaceHolder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: this.onExtensionDispose
      } as IPlaceholderCreateContentOptions);

      if (!this.topPlaceHolder) {
        // Failed to create PlacerHolder
        console.log("Failed to get Top PlaceHolder!");
        return;
      }

      if (this.topPlaceHolder.domElement) {
        let message: string = this.properties.title + " : "
          + this.properties.url;

        this.topPlaceHolder.domElement.innerHTML = `
          <div>
            <h2>${ message}</h2>
          </div>
        `;
      }
    }
  }

  private onExtensionDispose(): void {
    console.log("CoursePageApplicationCustomizer - onDispose Fired!");
  }
}
