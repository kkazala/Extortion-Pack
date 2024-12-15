import * as React from "react";
import { APIPermissions, APIPermissionsDef, UserInfo } from "../dal/types";
import { Body1, Body1Strong, Link, Spinner, Subtitle2, Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from "@fluentui/react-components";
import useStyles from "./styles";
import graphClient from "../dal/graphClient";
import { GraphFI } from "@pnp/graph";
import { PropertyContext } from "./SPFxSecurity";
import { APIPermissionsDefinitions } from "../dal/APIPermissions";

export type APIGraphUsersProps = {
	apiPermissions: APIPermissions;
	scopes: string[];
};

const APIGraphUsers = (props: APIGraphUsersProps): JSX.Element => {
	const { graphContext } = React.useContext<any>(PropertyContext);
	const { apiPermissions, scopes } = props;

	const styles = useStyles();
	const [users, setUsers] = React.useState<UserInfo[] | undefined>();
	const [assignedScopes, setAssignedScopes] = React.useState<APIPermissionsDef[]>([]);
	const [isLoading, setIsLoading] = React.useState<boolean>(true);

	React.useEffect(() => {
		const getContent = async (_graphContext: GraphFI): Promise<UserInfo[] | undefined> => {
			return await graphClient.getUsers(_graphContext);
		};

		if (apiPermissions.isUserRead) {
			getContent(graphContext)
				.then((users) => {
					setUsers(users);
					setIsLoading(false);
				})
				.catch((_) => {
					/** */
				});
		}

		const assignedScopesUser: APIPermissionsDef[] = APIPermissionsDefinitions.User.filter((_scope) => scopes.includes(_scope.value));
		setAssignedScopes(assignedScopesUser);
	}, []);

	return (
		<>
			{/* The following graph API permissions were evaluated */}
			<>
				<Subtitle2> The following graph API permissions are granted:</Subtitle2>
				<br />
				<br />
				{apiPermissions !== undefined && (
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell
									key="site"
									className={styles.narrowColumn}
								>
									API Permission
								</TableHeaderCell>
								<TableHeaderCell key="access">Description</TableHeaderCell>
							</TableRow>
						</TableHeader>
						{assignedScopes.length > 0 && (
							<TableBody>
								{assignedScopes.map((apiDef: APIPermissionsDef, index: number) => {
									return (
										<TableRow key={`api${index}`}>
											<TableCell>{apiDef.value}</TableCell>
											<TableCell>{apiDef.userConsentDescription}</TableCell>
										</TableRow>
									);
								})}
							</TableBody>
						)}
						{assignedScopes.length === 0 && (
							<TableRow key={`api0`}>
								<TableCell>--</TableCell>
								<TableCell>&nbsp;</TableCell>
							</TableRow>
						)}
					</Table>
				)}
				<br />
			</>
			{/* Introduction */}
			<>
				<Body1>
					Azure Active Directory (Azure AD) is Microsoft&apos;s cloud-based identity and access management service, used to manage and secure access to various resources, such as Microsoft
					365, Azure, etc.
					<br />
					Azure AD stores a wide range of information related to identities including{" "}
					<Link
						href="https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#json-representation"
						target="_blank"
					>
						user information
					</Link>
					, for example usernames and passwords, Email addresses, contact details (phone numbers, addresses), job titles and departments, etc.
					<br />
					<br />
					Access to this information is controlled through permissions granted to the application in Azure AD. Different permissions (e.g., User.Read, User.ReadBasic.All, User.ReadWrite.All)
					are required to access different pieces of user information, and the required permissions must be <Body1Strong>consented to by an administrator or the user themselves</Body1Strong>
					, depending on the scope of the data being accessed.
					<br />
					<br />
					<Body1Strong>Can malicious code change user&apos;s password?</Body1Strong>
					<br />
					To change a user&apos;s password using the Microsoft Graph API, <Body1Strong>User.ReadWrite.All or Directory.AccessAsUser.All</Body1Strong> are required. However, the user
					performing the password change <Body1Strong>must have the necessary permissions within Azure AD. Typically, an administrator</Body1Strong> would have these permissions.
					<br />
					<br />
					<Body1Strong>Can malicious code change user&apos;s profile properties?</Body1Strong>
					<br />
					For users whose source of authority is Windows Server Active Directory, you must use Windows Server Active Directory to update their identity, contact info, or job info.
					<br />
					Otherwise, blocking users from changing their profile properties can be achieved by setting up{" "}
					<Link
						href="https://learn.microsoft.com/en-us/exchange/permissions-exo/role-assignment-policies"
						target="_blank"
					>
						Role assignment policies in Exchange Online
					</Link>{" "}
					policies and permissions in Azure Active Directory.
				</Body1>
				<br />
				<br />
			</>
			{/* Why is it important? */}
			<>
				<Subtitle2>Why is it important?</Subtitle2>
				<br />
				<Body1>
					In the dark web marketplaces, personal data like emails and phone numbers are commonly traded commodities.
					<br />
					As of 2024, the average price for stolen email addresses ranges from <Body1Strong>$10 to $80</Body1Strong>, depending on the associated information and the email provider&apos;s
					popularity.
					<br />
					The value of such data on the dark web reflects its utility for cybercriminals, who can use it for various malicious activities such as phishing, identity theft, and unauthorized
					access to accounts. The continued demand for these items is driven by their potential to facilitate more profitable cybercrimes, like financial fraud and account takeovers.
					<br />
					<br />
					Even if your company publishes a list of all employees, not all the data available in Azure AD may be made public. Additionally, certain employees&apos; data might not be disclosed
					due to court orders aimed at ensuring their security. The exposure of their contact information could pose a serious risk.
				</Body1>
				<br />
				<br />
			</>
			{/* How can you stay secure? */}
			<>
				<Subtitle2>How can you stay secure?</Subtitle2>
				<br />
				<Body1>Remember that granting access to one SPFx solution effectively grants access to them all, also solutions that will be installed in the fututre.</Body1>
				<br />
				<br />
			</>
			{isLoading && <Spinner label="Loading your emails and calendars" />}
			{!isLoading && users !== undefined && (
				<>
					<Subtitle2>People that work with you (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="subject">Name</TableHeaderCell>
								<TableHeaderCell key="from">Info</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{users.map((value: UserInfo, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>{value.displayName}</TableCell>
									<TableCell>
										{value.mail}
										{value.businessPhones !== "" ? (
											<>
												<br />
												{value.businessPhones}
											</>
										) : (
											<></>
										)}
										{value.officeLocation !== "" ? (
											<>
												<br />
												{value.officeLocation}
											</>
										) : (
											<></>
										)}
									</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
				</>
			)}
		</>
	);
};
export default APIGraphUsers;
