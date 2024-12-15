import { Table, TableHeader, TableRow, TableHeaderCell, TableBody, Body1, Body1Strong, Subtitle2, TableCell, Spinner } from "@fluentui/react-components";
import * as React from "react";
import { ApimConfig, FileInfo, ListInfo, SiteInfo, SPSites } from "../dal/types";
import SiteDetails from "./SiteDetails";
import spClient_Http from "../dal/spHttmlClient";
import { HttpClient, SPHttpClient } from "@microsoft/sp-http";
import useStyles from "./styles";
import { SPFI } from "@pnp/sp";
import spfiClient from "../dal/spfiClient";
import { PropertyContext } from "./SPFxSecurity";

export type APIRestSitesProps = {
	onSitesLoaded: (sites: SiteInfo[]) => void;
	apimConfig?: ApimConfig;
};

const APIRestSites = (props: APIRestSitesProps): JSX.Element => {
	const { context, spfiContext } = React.useContext<any>(PropertyContext);

	const styles = useStyles();

	const [spSitesInfo, setSpSitesInfo] = React.useState<SiteInfo[] | undefined>();
	const [apimConfig, setApimConfig] = React.useState(props.apimConfig);
	const [isApimEnabled, setIsApimEnabled] = React.useState(false);
	const [isLoading, setIsLoading] = React.useState<boolean>(true);

	React.useEffect(() => {
		const getSites_spREST = async (_spHttpClient: SPHttpClient, _siteUrl: string): Promise<SPSites | undefined> => {
			//also includes user's MySite, will be removed because cannot be accessed via REST API
			const sites = await spClient_Http.getSites(_spHttpClient, _siteUrl);

			if (sites === undefined) return undefined; // exception

			return {
				sites: await spClient_Http.getSitesInfo(_spHttpClient, sites.sites),
			};
		};

		const getContent = async (_spHttpClient: SPHttpClient, _siteUrl: string): Promise<SPSites | undefined> => {
			return await getSites_spREST(_spHttpClient, _siteUrl);
		};

		getContent(context.spHttpClient, context.pageContext.site.absoluteUrl)
			.then((sitesRest) => {
				setSpSitesInfo(sitesRest?.sites);
				props.onSitesLoaded(sitesRest?.sites || []);
				setIsLoading(false);
			})
			.catch((_) => {
				/** */
			});
	}, []);
	React.useEffect(() => {
		if (props.apimConfig !== undefined) {
			setApimConfig(props.apimConfig);
			setIsApimEnabled(true);
		} else {
			setIsApimEnabled(false);
		}
	}, [props.apimConfig]);

	const onGetListsSP = async (_spHttpClient: SPHttpClient, siteRef: string): Promise<ListInfo[] | undefined> => {
		return await spClient_Http.getLists(_spHttpClient, siteRef);
	};
	const onShowListContentsSP = async (_spHttpClient: SPHttpClient, siteRef: string, listId: string, baseType: number): Promise<FileInfo[] | undefined> => {
		return await spClient_Http.getListContents(_spHttpClient, siteRef, listId, baseType);
	};
	const onSendListContentsSP = async (_spHttpClient: SPHttpClient, _spfiContext: SPFI, _httpClient: HttpClient, _apimConfig: ApimConfig, siteRef: string): Promise<void> => {
		await spfiClient.sendSiteContents(_spHttpClient, _spfiContext, _httpClient, _apimConfig, siteRef);
	};
	// have access to authentication cookies of a user who opens the SharePoint site.
	return (
		<>
			{/* The following API endpoint were evaluated */}
			<>
				<Subtitle2> The following API endpoints were used:</Subtitle2>
				<br />
				<br />
				<Table>
					<TableHeader className={styles.tableHeader}>
						<TableRow>
							<TableHeaderCell
								key="site"
								className={styles.narrowColumn}
							>
								API endpoint
							</TableHeaderCell>
							<TableHeaderCell key="access">Description</TableHeaderCell>
						</TableRow>
					</TableHeader>
					<TableBody>
						<TableRow key={`api1`}>
							{/*  always accessible */}
							<TableCell>SharePoint REST API</TableCell>
							<TableCell>SPFx solutions become part of SPO DNA and always have access to all SPO sites.</TableCell>
						</TableRow>
					</TableBody>
				</Table>
				<br />
			</>
			{/* Inroduction */}
			<>
				<Body1>
					Using this API, SharePoint Framework solutions have access to any SharePoint resources a user has access to, without requesting additional authentication or permissions. This
					approach is great for performance, because it doesn&apos;t cause additional overhead on runtime. However, it may{" "}
					<Body1Strong>expose you to potential security risks if you download Web Parts from unverified sources.</Body1Strong>
				</Body1>
				<br />
				When using this API from a context of SharePoint Online site, the solution doesn&apos;t have access to the contents of user&apos;s OneDrive.
				<br />
				<br />
			</>
			{/* Why is it important? */}
			<>
				<Subtitle2>Why is it important?</Subtitle2>
				<br />
				<Body1>
					Stolen files represent a significant threat as they can be utilized for financial gain, competitive advantage, extortion, social engineering, public exposure, etc.
					<ul>
						<li>Proprietary business information, trade secrets, and intellectual property can be sold to competitors or used to gain a competitive edge in the market.</li>
						<li>Sensitive operational data can be altered or deleted to disrupt business processes, causing financial and reputational damage.</li>
						<li>
							Information gleaned from stolen files can be used to craft convincing phishing emails, tricking recipients into revealing further sensitive information or downloading
							malware.
						</li>
						<li>
							Stolen files can be released publicly, causing reputational damage and legal issues, especially if they contain sensitive personal data or confidential business
							information.
						</li>
					</ul>
				</Body1>
			</>
			{/* How can you stay secure */}
			<>
				<Subtitle2>How can you stay secure?</Subtitle2>
				<br />
				<Body1>
					It is <Body1Strong>not possible to disable the SharePoint REST API access</Body1Strong>, and it is nearly impossible to review the code of a packaged solution.
				</Body1>
				<br />
				<br />
			</>
			{isLoading && <Spinner label="Loading your SharePoint sites" />}
			{!isLoading && spSitesInfo !== undefined && spSitesInfo.length > 0 && (
				<>
					<Subtitle2>Example SharePoint sites accessible with REST API (max 10)</Subtitle2>
					<br />
					<br />
					<Table size="small">
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="site">Site</TableHeaderCell>
								<TableHeaderCell
									key="access"
									className={styles.narrowColumn}
								>
									Show contents
								</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{spSitesInfo.map((value: SiteInfo, index: number) => (
								<SiteDetails
									key={`s${index}`}
									index={index}
									isApimEnabled={isApimEnabled}
									value={value}
									getLists={() => onGetListsSP(context.spHttpClient, value.url)}
									showContents={(listId: string, baseType: number) => onShowListContentsSP(context.spHttpClient, value.url, listId, baseType)}
									{...(apimConfig !== undefined && {
										sendContents: () => onSendListContentsSP(context.spHttpClient, spfiContext, context.httpClient, apimConfig, value.url),
									})}
								/>
							))}
						</TableBody>
					</Table>
				</>
			)}
		</>
	);
};

export default APIRestSites;
