import { Version } from '@microsoft/sp-core-library';
import {
IPropertyPaneConfiguration,
PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GreetingWebPart.module.scss';
import * as strings from 'GreetingWebPartStrings';

export interface IGreetingWebPartProps {
message: string;
}

export default class GreetingWebPart extends BaseClientSideWebPart <IGreetingWebPartProps> {

    public render(): void {
    this.domElement.innerHTML = `
    <div class="${ styles.greeting }">
        <div class="${ styles.container }">
            <div class="${ styles.row }">
                <div class="${ styles.column }">
                    <span class="${ styles.title }">GREETING</span>
                    <p class="${ styles.subTitle }">This is not my first SPx Web Parts.</p>
                    <p class="${ styles.description }">${escape(this.properties.message)}</p>
                </div>
            </div>
        </div>
    </div>`;
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
    PropertyPaneTextField('message', {
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
