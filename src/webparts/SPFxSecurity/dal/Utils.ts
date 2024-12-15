import { WebPartContext } from "@microsoft/sp-webpart-base";

export default class Utils {
	public static async GetAccessToken_MSGraph(context: WebPartContext): Promise<string> {
		const tokenProvider = await context.aadTokenProviderFactory.getTokenProvider();
		return await tokenProvider.getToken("https://graph.microsoft.com");
	}

	public static GetScopes_MSGraph(authToken: string, audience: string): string[] {
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

	public static ParsePermissions(scopes: string[]): {
		isSitesReadAll: boolean;
		isSitesSelected: boolean;
		isFilesRead: boolean;
		isFilesReadAll: boolean;
		isCalendarsRead: boolean;
		isCalendarsReadBasic: boolean;
		isMailRead: boolean;
		isMailReadBasic: boolean;
		isUserRead: boolean;
	} {
		const permissionsSites = ["Sites.FullControl.All", "Sites.Manage.All", "Sites.Read.All", "Sites.ReadWrite.All"];
		const permissionsFiles = ["Files.Read", "Files.ReadWrite"];
		const permissionsFilesAll = ["Files.Read.All", "Files.ReadWrite.All"];
		const permissionsCalendars = ["Calendars.Read", "Calendars.Read.Shared", "Calendars.ReadBasic", "Calendars.ReadWrite", "Calendars.ReadWrite.Shared"];
		const permissionsMail = ["Mail.Read", "Mail.Read.Shared", "Mail.ReadBasic", "Mail.ReadBasic.Shared", "Mail.ReadWrite", "Mail.ReadWrite.Shared"];
		const permissionsMailBasic = ["Mail.ReadBasic", "Mail.ReadBasic.Shared"];
		const permissionsUser = ["User.Read", "User.Read.All", "User.ReadBasic.All", "User.ReadWrite", "User.ReadWrite.All"];
		//Microsoft Graph Sites.Read.All
		return {
			isSitesReadAll: scopes.map((scope) => permissionsSites.includes(scope)).includes(true),
			isSitesSelected: scopes.includes("Sites.Selected"),

			isFilesRead: scopes.map((scope) => permissionsFiles.includes(scope)).includes(true),
			isFilesReadAll: scopes.map((scope) => permissionsFilesAll.includes(scope)).includes(true),

			isCalendarsRead: scopes.map((scope) => permissionsCalendars.includes(scope)).includes(true),
			isCalendarsReadBasic: scopes.includes("Calendars.ReadBasic"),

			isMailRead: scopes.map((scope) => permissionsMail.includes(scope)).includes(true),
			isMailReadBasic: scopes.map((scope) => permissionsMailBasic.includes(scope)).includes(true),
			isUserRead: scopes.map((scope) => permissionsUser.includes(scope)).includes(true),
		};
	}
}
