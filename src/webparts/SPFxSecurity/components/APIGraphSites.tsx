import { GraphFI } from "@pnp/graph";
import graphClient from "../dal/graphClient";
import * as React from "react";
import { APIPermissions, APIPermissionsDef, FileInfo, ListInfo, OneDriveContents, SiteInfo } from "../dal/types";
import { Table, TableHeader, TableRow, TableHeaderCell, TableBody, Body1, Body1Strong, Subtitle2, TableCell, Link, CounterBadge, Spinner } from "@fluentui/react-components";
import SiteDetails from "./SiteDetails";
import useStyles from "./styles";
import { PropertyContext } from "./SPFxSecurity";
import { APIPermissionsDefinitions } from "../dal/APIPermissions";
import { union } from "lodash";

export type APIGraphSitesProps = {
	apiSitesInfo: SiteInfo[] | undefined; //may be null if no API permissions
	apiPermissions: APIPermissions;
	scopes: string[];
};

const APIGraphSites = (props: APIGraphSitesProps): JSX.Element => {
	const { graphContext } = React.useContext<any>(PropertyContext);
	const { apiSitesInfo, apiPermissions, scopes } = props;

	const styles = useStyles();

	const [recentFiles, setRecentFiles] = React.useState<FileInfo[] | undefined>();
	const [sharedWithMe, setSharedFiles] = React.useState<FileInfo[] | undefined>();
	const [oneDriveContents, setOneDriveContents] = React.useState<OneDriveContents[] | undefined>();
	const [graphSitesInfo, setGraphSitesInfo] = React.useState<SiteInfo[] | undefined>();
	const [assignedScopes, setAssignedScopes] = React.useState<APIPermissionsDef[]>([]);
	const [isLoading, setIsLoading] = React.useState<boolean>(true);

	React.useEffect(() => {
		const getContent = async (_graphContext: GraphFI): Promise<{ recentFiles: FileInfo[] | undefined; sharedWithMe: FileInfo[] | undefined; oneDrive: OneDriveContents[] | undefined }> => {
			const [recentFiles, sharedWithMe, oneDrive] = await Promise.all([
				apiPermissions.isFilesRead || apiPermissions.isFilesReadAll ? graphClient.getOneDriveRecentFiles(_graphContext) : [], //recentFiles
				apiPermissions.isFilesReadAll ? graphClient.getFilesSharedWithMe(_graphContext) : [], //sharedWithMe
				apiPermissions.isFilesRead || apiPermissions.isFilesReadAll ? graphClient.getOneDriveContents(_graphContext) : [], //oneDrive
			]);

			return { recentFiles, sharedWithMe, oneDrive };
		};

		getContent(graphContext)
			.then(({ recentFiles, sharedWithMe, oneDrive }) => {
				setRecentFiles(recentFiles);
				setSharedFiles(sharedWithMe);
				setOneDriveContents(oneDrive);
				setIsLoading(false);
			})
			.catch((_) => {
				/** */
			});

		const assignedScopesSites: APIPermissionsDef[] = APIPermissionsDefinitions.Sites.filter((_scope) => scopes.includes(_scope.value));
		const assignedScopesFiles: APIPermissionsDef[] = APIPermissionsDefinitions.Files.filter((_scope) => scopes.includes(_scope.value));
		setAssignedScopes(union(assignedScopesSites, assignedScopesFiles));
	}, []);

	React.useEffect(() => {
		const getSites_spGraph = async (_graphContext: GraphFI): Promise<SiteInfo[] | undefined> => {
			//if Sites.Read.All not granted, returns [] to show the request completed and returned no results
			const sites: SiteInfo[] | undefined = await graphClient.getSites(_graphContext);
			return sites !== undefined ? await graphClient.getSitesInfo(_graphContext, sites) : undefined;
		};

		if (apiPermissions.isSitesReadAll) {
			getSites_spGraph(graphContext)
				.then((sites) => {
					setGraphSitesInfo(sites);
				})
				.catch((_) => {
					/** */
				});
		} else if (apiPermissions.isSitesSelected && apiSitesInfo !== undefined && apiSitesInfo.length > 0) {
			graphClient
				.getSitesInfo(graphContext, apiSitesInfo)
				.then((result: SiteInfo[]) => {
					setGraphSitesInfo(result);
				})
				.catch(() => {
					/**/
				});
		}
	}, [apiPermissions.isSitesReadAll, apiPermissions.isSitesSelected, apiSitesInfo]);

	const onGetListsGraph = async (_graphContext: GraphFI, siteRef: string): Promise<ListInfo[] | undefined> => {
		return await graphClient.getLists(_graphContext, siteRef);
	};
	const onShowListContentsGraph = async (_graphContext: GraphFI, siteRef: string, listId: string, baseType: number): Promise<FileInfo[] | undefined> => {
		return await graphClient.getListContents(_graphContext, siteRef, listId, baseType);
	};

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
					This{" "}
					<Link
						href="https://learn.microsoft.com/en-us/onedrive/developer/rest-api/?view=odsp-graph-online"
						target="_blank"
					>
						API
					</Link>{" "}
					can access all SharePoint Online sites the user has access to, along with content of their OneDrive. It may also access files shared with the user, and a list of recent files the
					user was working on.
					<br />
					The API Permissions model of SharePoint and OneDrive is{" "}
					<Link
						href="https://learn.microsoft.com/en-us/graph/permissions-selected-overview?tabs=http"
						target="_blank"
					>
						unique
					</Link>{" "}
					compared to other API endpoints, allowing for permissions to be granted for a specific site, list, or list item.
					<br />
					<br />
					<Body1Strong>Permissions to access user files</Body1Strong>
					<br />
					The <Body1Strong>Files.Read.All</Body1Strong> permission allows an app to read all files the signed-in user can access, while
					<br />
					<Body1Strong>Files.Read</Body1Strong> allows the app to read the signed-in user&apos;s files.
					<br />
					Essentially, <Body1Strong>Files.Read.All</Body1Strong> provides broader access to files across the organization, whereas <Body1Strong>Files.Read</Body1Strong> is limited to the
					user&apos;s own files, including those shared with the user.
					<br />
					The new <Body1Strong>Files.SelectedOperations.Selected</Body1Strong> manages application access at the file or library folder level, and requires an explicit{" "}
					<Link
						href="https://learn.microsoft.com/en-us/graph/permissions-selected-overview?tabs=http#roles"
						target="_blank"
					>
						role
					</Link>{" "}
					assignment to define which actions may be executed.
					<br />
					<br />
					<Body1Strong>Permissions to access SharePoint Online sites</Body1Strong>
					<br />
					While <Body1Strong>Sites.Read.All</Body1Strong> allows the application to read documents and list items in all site collections on behalf of the signed-in user, the <br />
					<Body1Strong>Sites.Selected</Body1Strong>
					ensures more granular access. By itself, <Body1Strong>Sites.Selected</Body1Strong> has no effect and role assignment must be configured by an administrator for the application to
					gain access to a site.
					<br />
					<br />
					In April 2024, Microsoft introduced even more{" "}
					<Link
						href="https://learn.microsoft.com/en-us/graph/permissions-selected-overview"
						target="_blank"
					>
						granular API permission levels
					</Link>{" "}
					supporting both, delegated and application, permission modes:
					<ul>
						<li>Lists.SelectedOperations.Selected,</li>
						<li>ListItems.SelectedOperations.Select,</li>
						<li>Files.SelectedOperations.Selected.</li>
					</ul>
				</Body1>
			</>
			{/* Why is it important? */}
			<>
				<Subtitle2>Why is it important?</Subtitle2>
				<br />
				<Body1>
					For the risks associated with data exfiltration, see the &quot;SharePoint REST API&quot; tab.
					<br />
					Compared to SharePoint REST API, Microsoft Graph additionally exposes OneDrive contents. If you use it to occasionally store personal files, you may be open to additional threats.
					Personal information such as social security numbers, addresses, and birth dates found in stolen files can be used to create false identities, apply for credit cards, loans, and
					other financial products.
				</Body1>
				<br />
				<br />
			</>
			{/* How can you stay secure? */}
			<>
				<Subtitle2>How can you stay secure?</Subtitle2>
				<br />
				<Body1>
					Always carefully review the requested permissions to ensure they are justified within the context of the app&apos;s functionality.{" "}
					<Body1Strong> Consider whether the potential productivity improvement outweigh the potential risks associated with granting excessive permissions.</Body1Strong>
					<br />
					<br />
					Aim for the lowest possible permission level and whenever possible, use fine-grained permissions.
					<br />
					Solutions request permissions to streamline and simplify the process, but you can grant and adjust them directly from Azure Portal. It&apos;s possible that a solution is requesting
					Sites.Read.All even though it may work with Sites.Selected. Contact the solution provider to evaluate if granting less permissions will be supported.
					<br />
				</Body1>
				<br />
				<br />
			</>
			{isLoading && <Spinner label="Loading your SharePoint and OneDrive contents" />}
			{!isLoading && graphSitesInfo !== undefined && graphContext !== undefined && (
				<>
					<Subtitle2>SharePoint sites accessible with Graph API (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
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
							{graphSitesInfo.map((value: SiteInfo, index: number) => (
								<SiteDetails
									key={`s${index}`}
									index={index}
									value={value}
									getLists={() => onGetListsGraph(graphContext, value.id)}
									showContents={(listId: string, baseType: number) => onShowListContentsGraph(graphContext, value.id, listId, baseType)}
								/>
							))}
						</TableBody>
					</Table>
				</>
			)}
			{!isLoading && recentFiles !== undefined && (
				<>
					<br />
					<br />
					<Subtitle2>Recent files you&apos;ve been working on (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="file">File name</TableHeaderCell>

								<TableHeaderCell
									key="date"
									className={styles.narrowColumn}
								>
									Last modified
								</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{recentFiles.map((value: FileInfo, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>
										<Link
											href={value.webUrl}
											target="_blank"
										>
											{value.name}
										</Link>
									</TableCell>
									<TableCell>{value.lastModifiedDateTime}</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
				</>
			)}
			{!isLoading && oneDriveContents !== undefined && (
				<>
					<br />
					<br />
					<Subtitle2>Contents of your OneDrive (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="file">File/folder name</TableHeaderCell>

								<TableHeaderCell
									key="date"
									className={styles.narrowColumn}
								>
									Child Count
								</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{oneDriveContents.map((value: OneDriveContents, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>
										<Link
											href={value.webUrl}
											target="_blank"
										>
											{value.name}
										</Link>
									</TableCell>
									<TableCell>
										<CounterBadge
											count={value.childCount}
											appearance="filled"
											size="small"
											color="informative"
										/>
									</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
				</>
			)}
			{!isLoading && sharedWithMe !== undefined && (
				<>
					<br />
					<br />
					<Subtitle2>Files shared with you (max 10)</Subtitle2>
					<br />
					<br />
					<Table>
						<TableHeader className={styles.tableHeader}>
							<TableRow>
								<TableHeaderCell key="file">File name</TableHeaderCell>

								<TableHeaderCell
									key="date"
									className={styles.narrowColumn}
								>
									Last modified
								</TableHeaderCell>
							</TableRow>
						</TableHeader>
						<TableBody>
							{sharedWithMe.map((value: FileInfo, index: number) => (
								<TableRow key={`s${index}`}>
									<TableCell>
										<Link
											href={value.webUrl}
											target="_blank"
										>
											{value.name}
										</Link>
									</TableCell>
									<TableCell>{value.lastModifiedDateTime}</TableCell>
								</TableRow>
							))}
						</TableBody>
					</Table>
				</>
			)}
		</>
	);
};

export default APIGraphSites;
