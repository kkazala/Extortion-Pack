import * as msal from "@azure/msal-browser";
import { IPublicClientApplication } from "@azure/msal-browser";
import { flatten, uniq } from "@microsoft/sp-lodash-subset";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { PnPClientStorage, PnPClientStorageWrapper } from "@pnp/core";
import { union } from "lodash";

export type KVPair<T> = {
	[key: string]: T;
};
export type TokenInfo = {
	homeAccountId: string;
	credentialType: string;
	secret: string;
	expiresOn: string;
	environment: string;
	clientId: string;
	realm: string;
	target: string;
	tokenType: string;
};

const AdminRolesMap: KVPair<string> = {
	"9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3": "Application Administrator", //Application Administrator [P]: Can create and manage all aspects of app registrations and enterprise apps. This role also grants the ability to consent for delegated permissions and application permissions, with the exception of application permissions for Azure AD Graph and Microsoft Graph.
	"cf1c38e5-3621-4004-a7cb-879624dced7c": "Application Developer", //Application Developer [P]: Can create application registrations independent of the 'Users can register applications' setting.
	"c4e39bd9-1100-46d3-8c65-fb160da0071f": "Authentication Administrator", //Authentication Administrator [P]: Can access to view, set and reset authentication method information for any non-admin user.
	"158c047a-c907-4556-b7ef-446551a6b5f7": "Cloud Application Administrator", //Cloud Application Administrator : Can create and manage all aspects of app registrations and enterprise apps except App Proxy.
	"62e90394-69f5-4237-9190-012177145e10": "Global Administrator", //Global Administrator [P]: Can manage access to all administrative features in Azure AD and other Microsoft services.
	"11648597-926c-4cf3-9c36-bcebb0ba8dcc": "Power Platform Administrator", //Power Platform Administrator: Can create and manage all aspects of Microsoft Dynamics 365, Power Apps and Power Automate.
};
//"b79fbf4d-3ef9-4689-8143-76b194e85509": exists in all non-guest accounts of the tenant. It does not refer to any administrator role,

export default class Utils {
	public static resourceGraph = "https://graph.microsoft.com";
	public static resourceAzure = "https://management.azure.com";

	public static GetAdminRoles(authToken: string): string[] {
		const tokenPartsRegex = /^([^.\s]*)\.([^.\s]+)\.([^.\s]*)$/;
		const matches = tokenPartsRegex.exec(authToken);
		const adminRoles: string[] = [];

		if (matches !== null) {
			const crackedToken = {
				header: matches[1],
				JWSPayload: matches[2],
				JWSSig: matches[3],
			};

			const decodedToken = {
				header: JSON.parse(atob(crackedToken.header)),
				payload: JSON.parse(atob(crackedToken.JWSPayload)),
				signature: crackedToken.JWSSig,
			};

			// console.log(decodedToken);
			if (Object.keys(decodedToken.payload).includes("wids")) {
				decodedToken.payload.wids
					.filter((wid: string) => wid !== "b79fbf4d-3ef9-4689-8143-76b194e85509") //exists in all non-guest accounts of the tenant
					.forEach((wid: string) => {
						adminRoles.push(AdminRolesMap[wid]);
					});
			}
		}
		return adminRoles;
	}
	//Get new Access token valid 1 hour from now
	public static async GetAccessToken(context: WebPartContext, resource: string): Promise<string> {
		try {
			const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
			return await tokenProvider.getToken(resource);
		} catch (error) {
			console.error(error);
			return "";
		}
	}

	//Get new Access token valid 1 hour from now
	public static async GetAccessToken_MSGraph(context: WebPartContext): Promise<string> {
		const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
		return await tokenProvider.getToken("https://graph.microsoft.com");
	}
	//Get new Access token valid 1 hour from now
	public static async GetAccessToken_Azure(context: WebPartContext): Promise<string> {
		const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
		return await tokenProvider.getToken("https://management.azure.com");
	}

	public static GetScopes(authToken: string, audience: string): string[] {
		const tokenPartsRegex = /^([^.\s]*)\.([^.\s]+)\.([^.\s]*)$/;
		const matches = tokenPartsRegex.exec(authToken);

		if (matches !== null) {
			const crackedToken = {
				header: matches[1],
				JWSPayload: matches[2],
				JWSSig: matches[3],
			};

			const decodedToken: { header: any; payload: any; signature: string } = {
				header: JSON.parse(atob(crackedToken.header)),
				payload: JSON.parse(atob(crackedToken.JWSPayload)),
				signature: crackedToken.JWSSig,
			};

			if (Object.keys(decodedToken.payload).includes("aud") && Object.keys(decodedToken.payload).includes("scp")) {
				if (decodedToken.payload.aud === audience) {
					return decodedToken.payload.scp.split(" ");
				}
			}
		}
		return [];
	}

	//Read tokens from local storage using MSAL
	public static async ReadTokensFromStorage(clientId: string, tenantId: string): Promise<TokenInfo[]> {
		const returnVals: TokenInfo[] = [];

		const msalInstance: IPublicClientApplication = await msal.createStandardPublicClientApplication({
			auth: {
				clientId: clientId,
				authority: `https://login.microsoftonline.com/${tenantId}`,
			},
			cache: {
				cacheLocation: "localStorage", // or "sessionStorage"
			},
		});

		const jsonToken = JSON.parse(JSON.stringify(msalInstance.getTokenCache()));
		Object.keys(jsonToken.storage.browserStorage).forEach((key) => {
			const value = jsonToken.storage.browserStorage[key]; //windowsStorage on Windows, probably other key on other OS
			Object.keys(value)
				.filter((k) => k.includes("login.windows.net"))
				.forEach((k) => {
					const info = JSON.parse(value[k]);
					if (info.credentialType === undefined) return;

					const expiresOn = new Date(parseInt(info.expiresOn) * 1000);
					//if expiresOn is in the past, skip
					if (expiresOn < new Date()) return;

					returnVals.push({
						homeAccountId: info.homeAccountId,
						credentialType: info.credentialType,
						secret: info.secret,
						expiresOn: expiresOn.toLocaleString(),
						environment: info.environment,
						clientId: info.clientId,
						realm: info.realm,
						target: info.target,
						tokenType: info.tokenType,
					});
				});
		});
		return returnVals;
	}

	//Read tokens from local storage using PnPClientStorage
	public static GetTokens(): TokenInfo[] {
		const storage = new PnPClientStorage();
		const allItems: Partial<PnPClientStorageWrapper> = storage.local;

		const returnVals: TokenInfo[] = [];

		Object.keys(allItems).forEach((key) => {
			if (key === "store") {
				const value = allItems[key as keyof PnPClientStorageWrapper];
				if (value !== undefined) {
					const values = Object.entries(value);
					values.forEach((val) => {
						if (val[0].includes("login.windows.net")) {
							const info = JSON.parse(val[1]);
							if (info.credentialType !== undefined) {
								const expiresOn = new Date(parseInt(info.expiresOn) * 1000);
								//if expiresOn is in the past, skip
								if (expiresOn < new Date()) {
									return;
								}
								returnVals.push({
									homeAccountId: info.homeAccountId,
									credentialType: info.credentialType,
									secret: info.secret,
									expiresOn: expiresOn.toLocaleString(),
									environment: info.environment,
									clientId: info.clientId,
									realm: info.realm,
									target: info.target,
									tokenType: info.tokenType,
								});
							}
						}
					});
				}
			}
		});
		return returnVals;
	}

	public static GetResources(allTokens: TokenInfo[]): string[] {
		const accessTokens: string[] = allTokens.filter((t) => t.credentialType === "AccessToken").map((t) => t.target);
		const urlRgx = /https?:\/\/[^\s]+/;
		const guidRgx = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/;

		const getMatch = (tokens: string[], rgx: RegExp): string[] => {
			return uniq(
				flatten(
					tokens
						.map((t) => {
							const m = t.match(rgx);
							return m ? m[0] : "";
						})
						.filter((t) => t !== null && t !== "")
				)
			);
		};

		const fromUrl: string[] = uniq(getMatch(accessTokens, urlRgx).map((t) => new URL(t).origin));
		const fromGuid: string[] = getMatch(accessTokens, guidRgx);

		return union(fromUrl, fromGuid);
	}
}
