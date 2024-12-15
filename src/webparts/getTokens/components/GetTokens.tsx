import * as React from "react";
import Utils from "../dal/Utils";
import type { GetTokensProps } from "./GetTokensProps";
import { InstructionsHeader, InstructionsUseCase1, InstructionsUseCase2, InstructionsUseCase3, PrintScopesAndRoles, PrintTokenInfo } from "./UtilsControls";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Button, Title3, Text, Body1Strong } from "@fluentui/react-components";
import { ArrowCounterclockwise32Regular, CheckmarkSquareRegular, DismissSquareRegular } from "@fluentui/react-icons";
import useStyles from "./styles";

export default function GetTokens(props: GetTokensProps): JSX.Element {
	const styles = useStyles();
	const [accessToken_Graph, setAccessToken_Graph] = React.useState<string>("");
	const [accessToken_AzureMgmnt, setAccessToken_AzureMgmnt] = React.useState<string>("");
	const [adminRoles, setAdminRoles] = React.useState<string[]>([]);
	const [scopesGraph, setScopesGraph] = React.useState<string[]>([]);
	const [scopesAzureMgmnt, setScopesAzureMgmnt] = React.useState<string[]>([]);

	const [hasRights1, setHasRights1] = React.useState<boolean>();
	const [hasRights2, setHasRights2] = React.useState<boolean>();
	const [hasRights3, setHasRights3] = React.useState<boolean>();

	/**  App registration and grant Application API permissions
	 * Role: Global Administrator
	 * API Permissions: `Application.ReadWrite.All`, `AppRoleAssignment.ReadWrite.All`
	 */

	/** Tenant takeover
	 *  Role: Global Administrator
	 * API Permissions:  `User.ReadWrite.All`, `RoleManagement.ReadWrite.Directory`
	 */

	/** Access / delete Azure resources
	 * Role: N/A
	 * API Permissions: Azure Service Management `user_impersonation`
	 */
	const GetTokens = async (context: WebPartContext): Promise<void> => {
		//#region graph
		//Get token for MS Graph
		const _msGraphToken = await Utils.GetAccessToken(context, Utils.resourceGraph);
		setAccessToken_Graph(_msGraphToken);

		//Get scopes form MS Graph
		const _scopesGraph = Utils.GetScopes(_msGraphToken, Utils.resourceGraph);
		setScopesGraph(_scopesGraph);

		//#endregion

		//#region Azure
		//Get token for Azure
		const _azureToken = await Utils.GetAccessToken(context, Utils.resourceAzure); //https://management.azure.com
		setAccessToken_AzureMgmnt(_azureToken);

		//Get scopes form Azure
		const _scopesAzure = Utils.GetScopes(_azureToken, Utils.resourceAzure);
		setScopesAzureMgmnt(_scopesAzure);
		//#endregion

		//Get Admin roles
		const adminRoles = Utils.GetAdminRoles(_msGraphToken);
		setAdminRoles(adminRoles);
	};

	React.useEffect(() => {
		GetTokens(props.context).catch((error) => {
			console.error(error);
		});
	}, []);

	return (
		<>
			{adminRoles && adminRoles.length > 0 && (
				<div>
					<Title3>Your Admin Roles</Title3>
					<ul>
						{adminRoles.map((role: string, index: number) => (
							<li key={index}>{role}</li>
						))}
					</ul>
				</div>
			)}
			{hasRights1 !== undefined && hasRights2 !== undefined && hasRights3 !== undefined && (
				<div>
					<Title3>Summary</Title3>
					<br />
					<Text>Based on your current role and API permissions assigned to the SharePoint service principal, you may execute the following use cases:</Text>
					<ul style={{ listStyleType: "none", marginLeft: "-20px" }}>
						<li>{hasRights1 ? <CheckmarkSquareRegular className={styles.icon} /> : <DismissSquareRegular className={styles.icon} />}Use Case 1: Tenant takeover </li>
						<li>
							{hasRights2 ? <CheckmarkSquareRegular className={styles.icon} /> : <DismissSquareRegular className={styles.icon} />}
							Use Case 2: Add App registration and grant Application API permissions{" "}
						</li>
						<li>{hasRights3 ? <CheckmarkSquareRegular className={styles.icon} /> : <DismissSquareRegular className={styles.icon} />}Use Case 3: Access/delete Azure resources </li>
					</ul>
					<Body1Strong className={styles.warning}>
						If you can execute these actions, a hacker may send the access token to their own server and execute these actions on your behalf.
					</Body1Strong>
					<br />
					<br />
				</div>
			)}
			<div>
				<Button
					appearance="primary"
					icon={<ArrowCounterclockwise32Regular />}
					onClick={() => GetTokens(props.context)}
				>
					Refresh Token
				</Button>
			</div>
			<div style={{ paddingTop: "20px" }}>
				<Title3>Use Case 1: Tenant takeover </Title3>
				<PrintScopesAndRoles
					resource={Utils.resourceGraph}
					scopes={{ assigned: scopesGraph, required: ["User.ReadWrite.All", "RoleManagement.ReadWrite.Directory"] }}
					roles={{ assigned: adminRoles, required: ["Global Administrator"] }}
					onScopesChecked={(result: boolean) => {
						setHasRights1(result);
					}}
				/>

				<>
					<InstructionsHeader text="create a new user and grant them Global Admin role" />
					<PrintTokenInfo tokenInfo={accessToken_Graph} />
					<InstructionsUseCase1
						userName={props.context.pageContext.user.displayName}
						userEmail={props.context.pageContext.user.loginName}
					/>
				</>
			</div>
			<div style={{ paddingTop: "20px" }}>
				<Title3>Use Case 2: Add App registration and grant Application API permissions </Title3>
				<PrintScopesAndRoles
					resource={Utils.resourceGraph}
					scopes={{ assigned: scopesGraph, required: ["Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"] }}
					roles={{ assigned: adminRoles, required: ["Global Administrator"] }}
					onScopesChecked={(result: boolean) => {
						setHasRights2(result);
					}}
				/>
				<>
					<InstructionsHeader text="create a new app registration and grant API permissions" />
					<PrintTokenInfo tokenInfo={accessToken_Graph} />
					<InstructionsUseCase2 userName={props.context.pageContext.user.displayName} />
				</>
			</div>
			<div style={{ paddingTop: "20px" }}>
				<Title3>Use Case 3: Access/delete Azure resources </Title3>
				<PrintScopesAndRoles
					resource={Utils.resourceAzure}
					scopes={{ assigned: scopesAzureMgmnt, required: ["user_impersonation"] }}
					roles={{ assigned: adminRoles, required: [] }}
					onScopesChecked={(result: boolean) => {
						setHasRights3(result);
					}}
				/>

				<>
					<InstructionsHeader text="sign in to Azure Management and create/delete resources" />
					<PrintTokenInfo tokenInfo={accessToken_AzureMgmnt} />
					<InstructionsUseCase3 />
				</>
			</div>
		</>
	);
}
