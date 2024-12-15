declare interface ISPFSecurityWebPartStrings {
	PropertyPaneDescription: string;
	BasicGroupName: string;
	URLFieldLabel: string;
	SubscriptionFieldLabel: string;
}

declare module "SPFSecurityWebPartStrings" {
	const strings: ISPFSecurityWebPartStrings;
	export = strings;
}
