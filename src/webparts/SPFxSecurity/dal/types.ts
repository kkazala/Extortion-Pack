import { HttpClient } from "@microsoft/sp-http";

export type SiteInfo = {
	name: string;
	url: string;
	id: string;
	canAccessLists?: boolean;
};
export type SPSites = {
	sites: SiteInfo[];
};
export type ListInfo = {
	id: string;
	title: string;
	itemCount: number;
	baseType: number; //0:list, 1: library
};
export type FileInfo = {
	name: string;
	lastModifiedDateTime: string;
	webUrl: string;
};
export type OneDriveContents = {
	name: string;
	webUrl: string;
	childCount: number;
};
export type ExoEvent = {
	subject: string;
	isAllDay: boolean;
	startEnd: string;
	body: string | undefined;
};

export type ExoMail = {
	subject: string;
	body: string | undefined;
	from: string;
};

export type UserInfo = {
	displayName: string;
	mail: string;
	officeLocation: string;
	businessPhones: string;
};
export type htmlClientExternalApiType = {
	httpClient: HttpClient;
};

//a subset of Microsoft Graph API permissions, used by the components
export type APIPermissions = {
	isSitesReadAll: boolean;
	isSitesSelected: boolean;
	isFilesRead: boolean;
	isFilesReadAll: boolean;
	isCalendarsRead: boolean;
	isCalendarsReadBasic: boolean;
	isMailRead: boolean;
	isMailReadBasic: boolean;
	isUserRead: boolean;
};

export type APIPermissionsDef = {
	adminConsentDescription: string;
	adminConsentDisplayName: string;
	type: string;
	userConsentDescription: string;
	userConsentDisplayName: string;
	value: string;
};

export type ApimConfig = {
	endpoint: string;
	key: string;
};
