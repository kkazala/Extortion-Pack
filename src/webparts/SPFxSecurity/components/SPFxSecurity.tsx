import * as React from "react";
import { FluentProvider, IdPrefixProvider, TabValue } from "@fluentui/react-components";
import APIRestSites from "./APIRestSites";
import APIGraphSites from "./APIGraphSites";
import APIGraphUsers from "./APIGraphUsers";
import APIGraphEXO from "./APIGraphEXO";
import { SPFxSecurityProps } from "./SPFxSecurityProps";
import AboutGovernance from "./AboutGovernance";
import TabListMenu from "./TabListMenu";
import { APIPermissions, SiteInfo } from "../dal/types";
import { HttpClient } from "@microsoft/sp-http";
import htmlClientExternalApi from "../dal/htmlClientExternalApi";
import Utils from "../dal/Utils";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx as spSPFx } from "@pnp/sp";
import { graphfi, SPFx as graphSPFx } from "@pnp/graph";

export const PropertyContext: any = React.createContext(undefined);

const SPFxSecurity = (props: SPFxSecurityProps): JSX.Element => {
	//#region State
	const _testValue = "testValue";

	//sites retrieved by APIRestSites component using restAPI. Used by APIGraphSites component is isSitesSelected is true and isSitesReadAll is false
	const [apiSitesInfo, setSpSitesInfo] = React.useState<SiteInfo[] | undefined>();

	//Api Permissions
	const [apiPermissions, setAPIPermissions] = React.useState<APIPermissions>();
	const [isPublicAPI, setIsPublicAPI] = React.useState<boolean | undefined>();
	//utils
	const [isLoading, setIsLoading] = React.useState<boolean>(true);
	const [isFetchingToken, setIsFetchingToken] = React.useState<boolean>(true);
	const [scopes, setScopes] = React.useState<string[]>([]);

	const [selectedValue, setSelectedValue] = React.useState<TabValue>("about");
	//#endregion

	React.useEffect(() => {
		const getContent = async (context: WebPartContext, _httpClient: HttpClient): Promise<void> => {
			//#region fetch fresh AccessToken
			const token = await Utils.GetAccessToken_MSGraph(context);
			const scopes = Utils.GetScopes_MSGraph(token, "https://graph.microsoft.com");
			setScopes(scopes);
			//#endregion

			//#region parse Graph permissions (susbset)
			const permissions = Utils.ParsePermissions(scopes);
			setAPIPermissions(permissions);
			//#endregion

			setIsFetchingToken(false);

			//#region ping 3rd party api with POST
			const sendResult = await htmlClientExternalApi.Send(_httpClient, _testValue);
			setIsPublicAPI(sendResult === _testValue);
			//#endregion

			setIsLoading(false);
		};

		// eslint-disable-next-line @typescript-eslint/no-floating-promises
		getContent(props.context, props.context.httpClient);
	}, []);

	return (
		<IdPrefixProvider value="SPFxSecurity">
			<FluentProvider id="SPFxSecurity">
				<TabListMenu
					onTabSelected={setSelectedValue}
					isLoading={isLoading}
				/>
				{apiPermissions !== undefined && isFetchingToken === false && (
					<>
						<PropertyContext.Provider
							value={{
								context: props.context,
								spfiContext: spfi().using(spSPFx(props.context)),
								graphContext: graphfi().using(graphSPFx(props.context)),
							}}
						>
							{selectedValue === "about" && (
								<AboutGovernance
									apiPermissions={apiPermissions}
									isFetchingToken={isFetchingToken}
									isPublicAPI={isPublicAPI}
								/>
							)}
							{selectedValue === "restAPI" && (
								<APIRestSites
									onSitesLoaded={(sites: SiteInfo[]) => {
										if (!apiPermissions.isSitesReadAll && apiPermissions.isSitesSelected) {
											setSpSitesInfo(sites);
										}
									}}
									apimConfig={props.apimConfig}
								/>
							)}
							{selectedValue === "graphAPISites" && (
								<APIGraphSites
									apiSitesInfo={apiSitesInfo}
									apiPermissions={apiPermissions}
									scopes={scopes}
								/>
							)}
							{selectedValue === "graphAPIEXO" && (
								<APIGraphEXO
									apiPermissions={apiPermissions}
									scopes={scopes}
								/>
							)}
							{selectedValue === "graphAPIUsers" && (
								<APIGraphUsers
									apiPermissions={apiPermissions}
									scopes={scopes}
								/>
							)}
						</PropertyContext.Provider>
					</>
				)}
			</FluentProvider>
		</IdPrefixProvider>
	);
};

export default SPFxSecurity;
