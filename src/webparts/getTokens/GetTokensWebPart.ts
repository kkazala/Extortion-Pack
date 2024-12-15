import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import * as React from "react";
import * as ReactDom from "react-dom";

import * as strings from "GetTokensWebPartStrings";
import GetTokens from "./components/GetTokens";
import { GetTokensProps } from "./components/GetTokensProps";

export interface IGetTokensWebPartProps {
	apimEndpoint: string;
	subscriptionKey: string;
}

export default class GetTokensWebPart extends BaseClientSideWebPart<IGetTokensWebPartProps> {
	public render(): void {
		const element: React.ReactElement<GetTokensProps> = React.createElement(GetTokens, {
			context: this.context,
			apimConfig: !!this.properties.apimEndpoint && !!this.properties.subscriptionKey ? { endpoint: this.properties.apimEndpoint, key: this.properties.subscriptionKey } : undefined,
		});

		ReactDom.render(element, this.domElement);
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse("1.0");
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: strings.PropertyPaneDescription,
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField("postURL", {
									label: strings.URLFieldLabel,
								}),
								PropertyPaneTextField("key", {
									label: strings.SubscriptionFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
