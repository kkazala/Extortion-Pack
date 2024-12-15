import { Body1, Body1Strong, Link, MessageBar, MessageBarBody, MessageBarTitle, Subtitle1, Subtitle2, Text, Tooltip, makeStyles, mergeClasses, tokens } from "@fluentui/react-components";
import * as React from "react";
import { APIPermissions } from "../dal/types";
import { ShieldError, ShieldOK } from "./Icons";

const useStyles = makeStyles({
	moreInfo: {
		backgroundColor: tokens.colorNeutralBackground1Hover,
		padding: "20px",
	},
	imgResults: {
		display: "block",
		position: "relative",
		width: "fit-content",
		margin: "auto",
	},
	chevron: {
		fontSize: "32px",
		verticalAlign: "bottom",
	},
	icons: {
		position: "absolute",
		fontSize: "24px",
		top: "0px",
	},
	iconOK: {
		color: tokens.colorPaletteGreenForeground1,
	},
	iconWarn: {
		color: tokens.colorPaletteDarkOrangeForeground1,
	},
	posFiles: {
		left: "198px",
	},
	posEvents: {
		left: "267px",
	},
	posEmails: {
		left: "342px",
	},
	posUsers: {
		left: "427px",
	},
	posAPI: {
		left: "518px",
		top: "32px",
		fontSize: "32px",
	},
});
export type AboutGovernanceProps = {
	isFetchingToken: boolean;
	apiPermissions: APIPermissions;
	isPublicAPI: boolean | undefined;
};

const AboutGovernance = (props: AboutGovernanceProps): JSX.Element => {
	const { isFetchingToken, apiPermissions, isPublicAPI } = props;
	const styles = useStyles();

	const _styleFilesOK = mergeClasses(styles.icons, styles.posFiles, styles.iconOK);
	const _styleFilesWarn = mergeClasses(styles.icons, styles.posFiles, styles.iconWarn);
	const _styleEventsOK = mergeClasses(styles.icons, styles.posEvents, styles.iconOK);
	const _styleEventsWarn = mergeClasses(styles.icons, styles.posEvents, styles.iconWarn);
	const _styleEmailsOK = mergeClasses(styles.icons, styles.posEmails, styles.iconOK);
	const _styleEmailsWarn = mergeClasses(styles.icons, styles.posEmails, styles.iconWarn);
	const _styleUsersOK = mergeClasses(styles.icons, styles.posUsers, styles.iconOK);
	const _styleUsersWarn = mergeClasses(styles.icons, styles.posUsers, styles.iconWarn);
	const _styleAPIOK = mergeClasses(styles.icons, styles.posAPI, styles.iconOK);
	const _styleAPIWarn = mergeClasses(styles.icons, styles.posAPI, styles.iconWarn);

	return (
		<>
			{!isFetchingToken && (
				<>
					<MessageBar
						key="success"
						intent="success"
					>
						<MessageBarBody>
							<MessageBarTitle>Meanwhile, this solution read claims in your Access Token, to find out which resources you have access to.</MessageBarTitle>
							Use the tabs above to see which information is{" "}
							<Body1Strong>
								accessible to{" "}
								<Text
									underline
									weight="bold"
								>
									all
								</Text>{" "}
								SPFx solutions in your tenant.
							</Body1Strong>
						</MessageBarBody>
					</MessageBar>
					<br />
				</>
			)}
			{/* Intro */}
			<Body1>
				When evaluating the security posture of your M365 and Azure environments in the context of SharePoint solutions installed in your tenant, consider the following key questions:
				<ul>
					<li>What information do these SPFx solutions have access to?</li>
					<li>How much control do you have over their access?</li>
					<li>Can this information be sent outside of your company without your consent?</li>
				</ul>
				It turns out the answers are: <Body1Strong>a lot, not much, and certainly yes.</Body1Strong>
				<br />
				<br />
				SPFx solutions, like this Web Part, may be used in SharePoint, Microsoft Teams, Office, Outlook, and Microsoft Viva. They can aggregate information from the above services, and any
				other API provided by Microsoft Graph or Azure.
				<br />
				<br />
				The applications are hosted in Sharepoint and run in the <Body1Strong>context of the current user</Body1Strong>, enabling communication between Microsoft services through ready-to-use
				libraries. TIt allows creating <Body1Strong>powerful solutions that deliver significant business value</Body1Strong>, allowing users to complete their tasks without needing to switch
				contexts or search for necessary information.
				<br />
				SPFx solutions may only request delegated permissions, which means they fully respect the user&apos;s privileges and do not access data or actions beyond their authorized scope.
				<br />
				<Body1Strong>However, they could also could become a potential attack vector...</Body1Strong>
				<br />
				<br />
			</Body1>
			{/* Safe by design? */}
			<Body1>
				<Subtitle1>Safe by design?</Subtitle1>
				<br />
				There are several types of{" "}
				<Tooltip
					content="Application Programming Interface"
					relationship="description"
				>
					<Text underline>APIs</Text>
				</Tooltip>{" "}
				that can be used by SPFx solutions: <Body1Strong>SharePoint REST API, Microsoft Graph, Azure and public APIs</Body1Strong>.
				<br />
				<br />
				<div className={styles.moreInfo}>
					<Link
						href="https://learn.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-sharepoint"
						target="_blank"
					>
						SharePoint REST API
					</Link>{" "}
					uses a <Body1Strong>built-in, ready-to-use authentication</Body1Strong> mechanism, which does{" "}
					<Body1Strong>
						NOT require additional permissions or approvals. All SPFx solutions may use it natively, reading all SharePoint and OneDrive for Business data the user has access to.
					</Body1Strong>
					<br />
					<br />
					<Link
						href="https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph"
						target="_blank"
					>
						Microsoft Graph
					</Link>{" "}
					provides access not only to SharePoint, but also other <Body1Strong>Microsoft 365 and Azure</Body1Strong> services.
					<br />
					APIs are called using <Body1Strong>SharePoint Online Client Extensibility</Body1Strong> service principal, acting on behalf of the user. The effective permissions are a result of
					the user&apos;s rights, and the API Permissions granted to the service principal.
					<br />
					<Body1Strong>
						These permissions apply to the{" "}
						<Link
							href="https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-aadhttpclient#granted-permissions-apply-to-the-entire-tenant"
							target="_blank"
						>
							entire tenant
						</Link>{" "}
						and can be used by any SPFx solution.
						<br />
						Malicious software can call various endpoints using permissions requested by other solutions. In case of an error, it may simply try another time.
					</Body1Strong>
					<br />
					<br />
					<Link
						href="https://learn.microsoft.com/en-us/sharepoint/dev/spfx/connect-to-anonymous-apis"
						target="_blank"
					>
						{" "}
						Public APIs
					</Link>{" "}
					may be called either anonymously, or by using function/API key or Basic Authentication. In this case, the current user&apos;s access authentication cookie is not available, but the
					credentials may be included in the solution itself.
				</div>
				<br />
			</Body1>
			{/* SPFx as a spyware */}
			<Body1>
				<Subtitle1>SPFx as a spyware</Subtitle1>
				<br />
				<Body1Strong>
					SPFx solutions, if created by malicious developers, may be used to gather information without users&apos; knowledge or consent. They can collect sensitive data and transmit it to
					external entities.
				</Body1Strong>
				<br />
				<br />
				Consider this scenario: Hackers or a state actor create an SPFx solution that offers substantial business value for your company. This solution could take various forms, such as an
				asset management system, a timesheet app, or even a QR code generator.
				<br />
				<br />
				It grows in popularity and soon, more and more people start using it.
				<br />
				However, without anyone&apos;s knowledge, this seemingly beneficial software secretly compromises security by transmitting sensitive data to unauthorized parties. And it does so in the
				context of every single person who opens a page with the malicious solution.
				<br />
				<br />
				<Body1Strong>
					The illustration below provides a summary of results when testing access to various types of information (such as email and documents) using different endpoints. This is likely
					just a subset.{" "}
					{isPublicAPI !== undefined &&
						(isPublicAPI ? <>This information MAY be sent outside of your company.</> : <>Great! It looks like access to public APIs is disabled in your organization.</>)}
				</Body1Strong>
				<br />
				<br />
				<Text align="center">
					<Text className={styles.imgResults}>
						<img
							width="603px"
							src={require("../assets/spywarev2.png")}
							alt="SPFx solution as a spyware, sends information to a 3rd party API"
						/>
						{apiPermissions.isSitesReadAll || apiPermissions.isSitesSelected ? <ShieldError className={_styleFilesWarn} /> : <ShieldOK className={_styleFilesOK} />}
						{apiPermissions.isCalendarsRead ? <ShieldError className={_styleEventsWarn} /> : <ShieldOK className={_styleEventsOK} />}
						{apiPermissions.isMailRead ? <ShieldError className={_styleEmailsWarn} /> : <ShieldOK className={_styleEmailsOK} />}
						{apiPermissions.isUserRead ? <ShieldError className={_styleUsersWarn} /> : <ShieldOK className={_styleUsersOK} />}
						{isPublicAPI !== undefined &&
							(isPublicAPI ? (
								<ShieldError
									className={_styleAPIWarn}
									text="Yes, all this information can be sent to an external endpoint"
								/>
							) : (
								<ShieldOK
									className={_styleAPIOK}
									text="Luckily, this information cannot be sent outside of your company."
								/>
							))}
					</Text>
				</Text>
				<br />
				<br />
				Review the content of other tabs in this Web Part, to find out which information is accessible by SPFx solutions, and what impact it may have on your organization.
			</Body1>
			{/* Staying safe */}
			<Body1>
				<br />
				<br />
				<Subtitle1>Staying safe</Subtitle1>
				<br />
				Data exfiltration is often a primary goal during cybersecurity attacks. Adversaries target specific organizations with the goal of accessing or stealing their confidential data while
				remaining undetected, either to resell it on the dark web or to post it for the world to see.
				<br />
				<br />
				<Subtitle2>Review API permissions</Subtitle2>
				<br />
				<Body1Strong>Regularly review API permissions</Body1Strong> assigned to the <Body1Strong>SharePoint Online Client Extensibility</Body1Strong> service principal, cross reference them
				with permissions requested by the SPFx solutions, and remove any unused permissions. The task may seem cumbersome, but you may use one of the scripts available in{" "}
				<Link
					href="https://pnp.github.io/script-samples/"
					target="_blank"
				>
					PnP Samples gallery
				</Link>
				, for example
				<ul>
					<li>
						<Link
							href="https://pnp.github.io/script-samples/spo-get-spfx-apipermissions/README.html?tabs=pnpps"
							target="_blank"
						>
							GET API Permissions for SPFx solutions
						</Link>
						, to get a summary of all SPFx extensions installed in SPO sites and API permissions assigned, or
					</li>
					<li>
						<Link
							href="https://pnp.github.io/script-samples/spo-delete-unused-spfx-apipermissions/README.html?tabs=graphps"
							target="_blank"
						>
							Remove unused API Permissions assigned to SharePoint Online Client Extensibility Web Application Principal
						</Link>{" "}
						to automatically remove unused API permissions assigned to the service principal.
					</li>
				</ul>
				<br />
				<Subtitle2>Use Application Insights to track external traffic</Subtitle2>
				<br />
				Although it is impossible to proactively block external traffic initiated by SPFx solutions, you may review the traffic generated by your SharePoint sites. This can be achieved using
				the{" "}
				<Link
					href="https://github.com/pnp/sp-dev-fx-webparts/tree/88310656fe9cf89e472afe1685bcf08e532971e2/samples/js-applicationinsights-api-calls-tracking"
					target="_blank"
				>
					{" "}
					SPFx Application Customizer
				</Link>{" "}
				deployed tenant-wide to all SPO sites, that tracks all API requests and sends them to Application Insights.
				<br />
				<br />
				This solution allows whitelisting &apos;safe&apos; endpoints with a goal of reducing the amount of data logged. It also allows for temporarily disabling logging to facilitate
				randomized &quot;hunting&quot; without continuously generating large volumes of data. Results can be reviewed using the Application Map in Application Insights or by executing KUSTO
				query against the Application Insights logs, offering powerful tools for analyzing and understanding API usage patterns.
			</Body1>
		</>
	);
};
export default AboutGovernance;
