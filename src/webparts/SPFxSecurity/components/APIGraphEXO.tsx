import * as React from "react";
import { APIPermissions, APIPermissionsDef, ExoEvent, ExoMail } from "../dal/types";
import { Body1, Body1Strong, Link, Spinner, Subtitle2, Table, TableBody, TableCell, TableHeader, TableHeaderCell, TableRow } from "@fluentui/react-components";
import useStyles from "./styles";
import { GraphFI } from "@pnp/graph";
import graphClient from "../dal/graphClient";
import { PropertyContext } from "./SPFxSecurity";
import { union } from "lodash";
import { APIPermissionsDefinitions } from "../dal/APIPermissions";

export type APIGraphEXOProps = {
	apiPermissions: APIPermissions;
	scopes: string[];
};

const APIGraphEXO = (props: APIGraphEXOProps): JSX.Element => {
	const { graphContext } = React.useContext<any>(PropertyContext);
	const { apiPermissions, scopes } = props;

	const styles = useStyles();
	const [events, setEvents] = React.useState<ExoEvent[] | undefined>();
	const [calendars, setCalendars] = React.useState<string[] | undefined>();
	const [mails, setMails] = React.useState<ExoMail[] | undefined>();
	const [assignedScopes, setAssignedScopes] = React.useState<APIPermissionsDef[]>([]);
	const [isLoading, setIsLoading] = React.useState<boolean>(true);

	React.useEffect(() => {
		const getContent = async (
			_graphContext: GraphFI,
			_apiPermissions: APIPermissions
		): Promise<{ _events: ExoEvent[] | undefined; _calendars: string[] | undefined; _mails: ExoMail[] | undefined }> => {
			const [_calendars, _events, _mails] = await Promise.all([
				_apiPermissions.isCalendarsRead ? graphClient.getAllCalendars(_graphContext) : [], //calendars
				_apiPermissions.isCalendarsRead ? graphClient.getEvents(_graphContext) : [], //events
				_apiPermissions.isMailRead && _apiPermissions.isUserRead ? graphClient.getEmails(_graphContext) : [], //mail
			]);

			return { _events, _calendars, _mails };
		};

		getContent(graphContext, apiPermissions)
			.then(({ _events, _calendars, _mails }) => {
				setEvents(_events);
				setCalendars(_calendars);
				setMails(_mails);
				setIsLoading(false);
			})
			.catch((_) => {
				/** */
			});

		const assignedScopesSites: APIPermissionsDef[] = APIPermissionsDefinitions.Calendars.filter((_scope) => scopes.includes(_scope.value));
		const assignedScopesFiles: APIPermissionsDef[] = APIPermissionsDefinitions.Mail.filter((_scope) => scopes.includes(_scope.value));
		const assignedScopesUser: APIPermissionsDef[] = APIPermissionsDefinitions.User.filter((_scope) => scopes.includes(_scope.value));
		setAssignedScopes(union(assignedScopesSites, assignedScopesFiles, assignedScopesUser));
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
					If your company uses <Body1Strong>Exchange Online</Body1Strong>, Microsoft Graph can be leveraged to access your emails and calendar.
				</Body1>
				<br />
				<br />
				<Body1Strong>Permissions to read Calendars</Body1Strong>
				<br />
				<Body1Strong>Calendars.ReadBasic</Body1Strong> permission allowing to read events in user calendars, excluding properties like body, attachments, and extensions, and the{" "}
				<Body1Strong>Calendars.Read</Body1Strong> which grants access to all event.
				<br />
				<Body1Strong>Calendars.Read.Shared </Body1Strong> allows the app to read events in all calendars that the user can access, including delegate and shared calendars.
				<br />
				<br />
				<Body1Strong>Permissions to read emails</Body1Strong>
				<br />
				<Body1Strong>Mail.ReadBasic</Body1Strong> permission allows an app to read basic information from user mailboxes, such as subject lines and sender details and <br />
				<Body1Strong>Mail.Read</Body1Strong> provides allows the app to read all mail messages in user mailboxes, including body content and attachments.
				<br />
				<Body1Strong>Mail.Read.Shared</Body1Strong> (delegated permissions only) allows to read mail a user can access, including their own and shared mail.
				<br />
				<Body1Strong>Mail.ReadBasic.Shared</Body1Strong> (delegated permissions only) enables the app to read mail the signed-in user can access, including their own and shared mail, except
				for body, bodyPreview, uniqueBody, attachments, extensions, and any extended properties.
				<br />
				<br />
			</>
			{/* Why is it important */}
			<>
				<Subtitle2>Why is it important?</Subtitle2>
				<br />
				<Body1>
					Stolen emails and calendar events can be valuable tools for hackers to exploit in various malicious way.
					<ul>
						<li>
							Hackers can use stolen email addresses to send deceptive emails that appear legitimate, tricking recipients into clicking malicious links or providing sensitive
							information. By having access to past emails, hackers can craft highly personalized and convincing phishing emails to target specific individuals, increasing the likelihood
							of success.
						</li>
						<li>
							Emails can contain sensitive information such as personal identification numbers, financial details, and confidential business information, which can be exploited for
							identity theft, financial fraud, or corporate espionage.
						</li>
						<li>
							Detailed calendar events can provide insights into an individual&aposs schedule, making it easier for hackers to impersonate colleagues or associates and execute social
							engineering attacks.
						</li>
						<li>Knowing the locations and times of meetings can facilitate physical security breaches, such as unauthorized access to secure areas or stalking.</li>
						<li>Calendar events can reveal business strategies, upcoming projects, or mergers and acquisitions, which can be valuable for corporate espionage.</li>
					</ul>
				</Body1>
			</>
			{/* How can you stay secure? */}
			<>
				<Subtitle2>How can you stay secure?</Subtitle2>
				<br />
				<Body1>
					Always carefully review the requested permissions to ensure they are justified within the context of the app&apos;s functionality.{" "}
					<Body1Strong> Consider whether the potential productivity improvement outweigh the potential risks associated with granting excessive permissions.</Body1Strong>
					<br />
					Aim for the lowest possible permission level. Although Graph API for Exchange does not offer granular API permissions (like API for SharePoint and OneDrive), you can decide how
					many details an app can access.
					<br />
					<Link href="https://learn.microsoft.com/en-us/exchange/permissions-exo/application-rbac">Role Based Access Control for Applications in Exchange Online</Link> and{" "}
					<Link href="https://learn.microsoft.com/en-us/graph/auth-limit-mailbox-access">ApplicationAccessPolicy</Link> allow limiting application permissions to specific Exchange Online
					mailboxes, but they work with application permissions only and have <Body1Strong>no effect on delegated permissions</Body1Strong>.
				</Body1>
				<br />
				<br />
			</>
			{isLoading && <Spinner label="Loading your emails and calendars" />}
			{!isLoading && calendars !== undefined && (
				<>
					<Subtitle2>Your calendars:</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="file">Calendar name</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{calendars.map((value: string, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>{value}</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
					<br />
					<br />
				</>
			)}
			{!isLoading && events !== undefined && (
				<>
					<Subtitle2>Events in your calendar (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="event">Event</TableHeaderCell>
								<TableHeaderCell key="time">Start - End</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{events.map((value: ExoEvent, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>
										{value.subject}{" "}
										{value.body !== undefined ? (
											<>
												<br />
												{value.body}
											</>
										) : (
											<></>
										)}
									</TableCell>
									<TableCell>{value.startEnd}</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
					<br />
					<br />
				</>
			)}
			{!isLoading && mails !== undefined && (
				<>
					<Subtitle2>Recent emails in your Inbox (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="subject">Subject</TableHeaderCell>
								<TableHeaderCell key="from">From</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{mails.map((value: ExoMail, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>
										{value.subject}{" "}
										{value.body !== undefined ? (
											<>
												<br />
												{value.body}
											</>
										) : (
											<></>
										)}
									</TableCell>
									<TableCell>{value.from}</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
				</>
			)}
		</>
	);
};
export default APIGraphEXO;
