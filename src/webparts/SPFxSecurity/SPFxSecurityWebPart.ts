import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { SPFxSecurityProps } from "./components/SPFxSecurityProps";
import SPFxSecurity from "./components/SPFxSecurity";
import * as strings from "SPFSecurityWebPartStrings";

export interface ISPFxSecurityWebPartProps {
	apimEndpoint: string;
	subscriptionKey: string;
}

export default class SPFxSecurityWebPart extends BaseClientSideWebPart<ISPFxSecurityWebPartProps> {
	public render(): void {
		const element: React.ReactElement<SPFxSecurityProps> = React.createElement(SPFxSecurity, {
			context: this.context,
			apimConfig: !!this.properties.apimEndpoint && !!this.properties.subscriptionKey ? { endpoint: this.properties.apimEndpoint, key: this.properties.subscriptionKey } : undefined,
		});
		ReactDom.render(element, this.domElement);
	}

	protected onInit(): Promise<void> {
		return Promise.resolve();
	}

	protected onDispose(): void {
		ReactDom.unmountComponentAtNode(this.domElement);
	}

	protected get dataVersion(): Version {
		return Version.parse(this.context.manifest.version);
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: {
						description: "APIM configuration",
					},
					groups: [
						{
							groupName: strings.BasicGroupName,
							groupFields: [
								PropertyPaneTextField("apimEndpoint", {
									label: strings.URLFieldLabel,
								}),
								PropertyPaneTextField("subscriptionKey", {
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
