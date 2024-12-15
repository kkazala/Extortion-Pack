import * as React from "react";
import { TokenInfo } from "../dal/Utils";
import { Table, TableBody, TableHeader, TableHeaderCell, TableRow, Text, Subtitle1, Button, Tooltip } from "@fluentui/react-components";
import useStyles from "./styles";
import { configInfo } from "../dal/types";
import { CopyFilled } from "@fluentui/react-icons";

const copyToClipboard = (text: string): void => {
	navigator.clipboard.writeText(text).catch((_) => {
		/**/
	});
};

function PrintScopesAndRoles(props: { resource: string; scopes: configInfo; roles: configInfo; onScopesChecked: (hasRights: boolean) => void }): JSX.Element {
	const styles = useStyles();

	React.useEffect(() => {
		const hasAPIPermissions = props.scopes.required.filter((scope) => props.scopes.assigned.includes(scope)).length === props.scopes.required.length;
		const hasAdminRoles = props.roles.required.length === 0 || props.roles.required.filter((role) => props.roles.assigned.includes(role)).length === props.roles.required.length;

		props.onScopesChecked(hasAPIPermissions && hasAdminRoles);
	}, [props.scopes.assigned, props.roles.assigned]);

	return (
		<>
			<Table className={styles.addMargin}>
				<TableHeader className={styles.tableHeader}>
					<TableRow>
						<TableHeaderCell key="site">Required API Permission for {props.resource}</TableHeaderCell>
						<TableHeaderCell
							key="access"
							className={styles.narrowColumn}
						>
							Granted?
						</TableHeaderCell>
					</TableRow>
				</TableHeader>
				<TableBody>
					{props.scopes.required.map((scope) => {
						return (
							<TableRow key={scope}>
								<td>{scope}</td>
								<td>{props.scopes.assigned.includes(scope) ? "yes" : "no"}</td>
							</TableRow>
						);
					})}
				</TableBody>
			</Table>
			{props.roles.required.length > 0 && (
				<Table className={styles.addMargin}>
					<TableHeader className={styles.tableHeader}>
						<TableRow>
							<TableHeaderCell key="site">Required admin role </TableHeaderCell>
							<TableHeaderCell
								key="access"
								className={styles.narrowColumn}
							>
								Granted?
							</TableHeaderCell>
						</TableRow>
					</TableHeader>
					<TableBody>
						{props.roles.required.map((role) => {
							return (
								<TableRow key={role}>
									<td>{role}</td>
									<td>{props.roles.assigned.includes(role) ? "yes" : "-"}</td>
								</TableRow>
							);
						})}
					</TableBody>
				</Table>
			)}
		</>
	);
}

function PrintTokenInfo(props: { tokenInfo: string }): JSX.Element {
	const tokenInfo = props.tokenInfo;
	const styles = useStyles();

	return (
		<>
			<Text style={{ display: "block", paddingTop: "20px" }}>
				1. Set the <Text font="monospace">$accesstoken</Text> to the following value:
			</Text>
			<Text style={{ display: "flex" }}>
				{/* display subsctring of 100 characters for tokenInfo */}

				<br />
				<Text style={{ textWrap: "wrap", wordWrap: "break-word", width: "90%" }}>{!!tokenInfo ? `${tokenInfo.substring(0, 100)}...` : "N/A"}</Text>
				<Tooltip
					content={{ children: "Copy access token to clipboard", className: styles.tooltip }}
					relationship="label"
				>
					<Button
						className={styles.btnCopy}
						icon={<CopyFilled className={styles.iconCopy} />}
						disabled={!!!tokenInfo}
						onClick={() => {
							copyToClipboard(tokenInfo);
						}}
					/>
				</Tooltip>
			</Text>
		</>
	);
}
function PrintTokenInfoMulti(props: { tokenInfo: TokenInfo[] }): JSX.Element {
	return (
		<>
			{props.tokenInfo.map((token: TokenInfo) => {
				return (
					<>
						<b>Valid until: {token.expiresOn}</b>
						<br />
						{token.target !== "" && (
							<>
								<b>{token.target}</b>
								<br />
							</>
						)}
						<PrintTokenInfo tokenInfo={token.secret} />
						<br />
						<br />
					</>
				);
			})}
		</>
	);
}
function InstructionsHeader(props: { text: string }): JSX.Element {
	return (
		<div>
			<Subtitle1>Instructions</Subtitle1>
			<Text style={{ display: "flex" }}>Use the script below to {props.text}</Text>
		</div>
	);
}

function InstructionsUseCase1(props: { userName: string; userEmail: string }): JSX.Element {
	const upn = `a${props.userEmail}`;
	return (
		<div style={{ paddingTop: "20px" }}>
			<Text>2. Execute the following PS script:</Text>
			<br />
			<div style={{ backgroundColor: "#fafafa", padding: "20px" }}>
				<Text
					font="monospace"
					wrap
				>
					# Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force <br />
					Import-Module Microsoft.Graph.Authentication
					<br />
					Import-Module Microsoft.Graph.Users
					<br />
					$secure=ConvertTo-SecureString $accessToken -AsPlainText -Force <br />
					Connect-MgGraph -AccessToken $secure <br />
					$ctx=Get-MgContext <br />
					<br />
					# create a new user <br />
					$PasswordProfile = @{"{"}Password = &quot;xWwvJ=6NMw+bWH-d&quot;{"}"}
					<br />
					$user= New-MgUser -DisplayName &apos;{props.userName}&apos; -PasswordProfile $PasswordProfile -AccountEnabled -MailNickName &apos;{props.userName.replace(" ", "")}&apos;
					-UserPrincipalName &apos;
					{upn}
					&apos;
					<br />
					<br />
					# assign permanent Global Admin role <br />
					$role = Get-MgDirectoryRole | Where-Object {"{"}$_.displayName -eq &apos;Global Administrator&apos;{"}"} <br />
					$newRoleMember =@{"{"}&quot;@odata.id&quot; = &quot;https://graph.microsoft.com/v1.0/users/$($user.Id)&quot; {"}"} <br />
					New-MgDirectoryRoleMemberByRef -DirectoryRoleId $role.Id -BodyParameter $newRoleMember
				</Text>
			</div>
		</div>
	);
}

function InstructionsUseCase2(props: { userName: string }): JSX.Element {
	const apiPermissionsIds = ["1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9", "06b708a9-e830-4db3-a914-8e69da51d44f", "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8", "741f803b-c850-494e-b5df-cde7c675a1ca"];

	return (
		<div style={{ paddingTop: "20px" }}>
			<Text>2. Execute the following PS script:</Text>
			<br />
			<div style={{ backgroundColor: "#fafafa", padding: "20px" }}>
				<Text
					font="monospace"
					wrap
				>
					# Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force <br />
					Import-Module Microsoft.Graph.Applications
					<br />
					$secure=ConvertTo-SecureString $accessToken -AsPlainText -Force <br />
					Connect-MgGraph -AccessToken $secure <br />
					$ctx=Get-MgContext <br />
					<br />
					# type=&quot;Role&quot; means App-only API permissions <br />
					$requiredResourceAccess = {"(@{"}
					<br />
					&emsp;&quot;resourceAccess&quot; ={"("}
					{apiPermissionsIds.map((id, index) => {
						return (
							<>
								<br />
								&emsp;@{"{"}
								<br />
								&emsp;&emsp;id = &quot;{id}&quot; <br />
								&emsp;&emsp;type = &quot;Role&quot;
								<br />
								&emsp;{"}"}
								{index < apiPermissionsIds.length - 1 ? "," : ""}
							</>
						);
					})}
					<br />
					&emsp;{")"}
					<br />
					&emsp;&quot;resourceAppId&quot; = &quot;00000003-0000-0000-c000-000000000000&quot; #MS Graph
					<br />
					{"})"}
					<br />
					<br />
					# create a new app registration <br />
					$app = New-MgApplication -DisplayName &quot;{props.userName.replace(" ", "")} App Registration&quot; -RequiredResourceAccess $requiredResourceAccess
					<br />
					<br />
					# grant admin consent
					<br />
					$graphSpId = $(Get-MgServicePrincipal -Filter &quot;appId eq &apos;00000003-0000-0000-c000-000000000000&apos;&quot;).Id <br />
					$sp = New-MgServicePrincipal -AppId $app.appId <br />
					{apiPermissionsIds.map((id) => {
						return (
							<>
								New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -AppRoleId &quot;{id}&quot; -ResourceId $graphSpId
								<br />
							</>
						);
					})}
					# create client secret
					<br /> $cred = Add-MgApplicationPassword -ApplicationId $app.id
				</Text>
			</div>
		</div>
	);
}

function InstructionsUseCase3(): JSX.Element {
	return (
		<div style={{ paddingTop: "20px" }}>
			<Text>2. Execute the following PS script:</Text>
			<br />
			<div style={{ backgroundColor: "#fafafa", padding: "20px" }}>
				<Text
					font="monospace"
					wrap
				>
					Connect-AzAccount -AccessToken $accessToken -AccountId &apos;https://management.azure.com/&apos; -TenantId $TenantId
					<br />
					#create a resource group <br />
					New-AzResourceGroup -Name exampleGroup -Location westus
				</Text>
			</div>
		</div>
	);
}

export { PrintTokenInfo, PrintTokenInfoMulti, PrintScopesAndRoles, InstructionsHeader, InstructionsUseCase1, InstructionsUseCase2, InstructionsUseCase3 };
