import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
    BaseClientSideWebPart,
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    IPropertyPaneDropdownOption,
    PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import WebhookSubscription from './components/WebhookSubscription';
import { IWebhookSubscriptionProps } from './components/IWebhookSubscriptionProps';
import { IWebhookSubscriptionWebPartProps, ISPLists, ISPList } from './IWebhookSubscriptionWebPartProps';
import { SPHttpClientResponse, SPHttpClient } from "@microsoft/sp-http";

export default class WebhookSubscriptionWebPart extends BaseClientSideWebPart<IWebhookSubscriptionWebPartProps> {
    private _listOptions: IPropertyPaneDropdownOption[] = [];

    public render(): void {
        const element: React.ReactElement<IWebhookSubscriptionProps> = React.createElement(
            WebhookSubscription,
            {
                listname: this.properties.listname || "",
                subscriptionUrl: this.properties.subscriptionUrl || "",
                clientState: this.properties.clientState,
                context: this.context
            }
        );

        ReactDom.render(element, this.domElement);
    }

    // Override onInit function
    protected onInit<T>(): Promise<T> {
        this._GetListDataAsync()
            .then((response) => {
                this._listOptions = response.value.map((list: ISPList) => {
                    // Map each of the lists to a dropdown option
                    return {
                        key: list.Title,
                        text: list.Title
                    };
                });
            });

        return Promise.resolve();
    }

    // Get all the lists from the current SharePoint site
    private _GetListDataAsync(): Promise<ISPLists> {
        return this.context.spHttpClient
            .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Title`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                return response.json();
            });
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: "Webhook web part configuration"
                    },
                    groups: [
                        {
                            groupName: "Settings",
                            groupFields: [
                                PropertyPaneDropdown('listname', {
                                    label: 'Choose the list for your subscriptions:',
                                    options: this._listOptions
                                }),
                                PropertyPaneTextField('subscriptionUrl', {
                                    label: 'Specify your subscription URL:'
                                }),
                                PropertyPaneTextField('clientState', {
                                    label: 'Client state for the subscription (optional):'
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }

    protected get disableReactivePropertyChanges() : boolean {
        return true;
    }
}
