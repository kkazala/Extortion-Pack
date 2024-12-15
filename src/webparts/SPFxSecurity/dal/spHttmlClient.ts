import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as moment from "moment";
import { SiteInfo, ListInfo, FileInfo, SPSites } from "./types";
import { sortBy } from "@microsoft/sp-lodash-subset";

export default class spClient_Http {
	public static async getSites(spHttpClient: SPHttpClient, url: string, rowlimit = 11): Promise<SPSites | undefined> {
		const parseSites = (res: any): SiteInfo[] => {
			const tenantName = url.replace(/^https?:\/\//, "");

			return res.PrimaryQueryResult.RelevantResults.Table.Rows.map((row: any) => {
				const siteId = row.Cells.find((cell: any) => cell.Key === "SiteId").Value;
				const webId = row.Cells.find((cell: any) => cell.Key === "WebId").Value;
				const graphId = `${tenantName},${siteId},${webId}`;
				return {
					name: row.Cells.find((cell: any) => cell.Key === "Title").Value,
					url: row.Cells.find((cell: any) => cell.Key === "Path").Value,
					id: graphId,
				};
			})
				.filter((r: SiteInfo) => r.url.indexOf("/personal/") === -1)
				.slice(0, 10);
		};
		const getMySite = (res: any): SiteInfo | undefined => {
			const tenantName = url.replace(/^https?:\/\//, "");

			return res.PrimaryQueryResult.RelevantResults.Table.Rows.map((row: any) => {
				const siteId = row.Cells.find((cell: any) => cell.Key === "SiteId").Value;
				const webId = row.Cells.find((cell: any) => cell.Key === "WebId").Value;
				const graphId = `${tenantName},${siteId},${webId}`;
				return {
					name: row.Cells.find((cell: any) => cell.Key === "Title").Value,
					url: row.Cells.find((cell: any) => cell.Key === "Path").Value,
					id: graphId,
				};
			}).filter((r: SiteInfo) => r.url.indexOf("/personal/") !== -1)?.[0];
		};

		return spHttpClient
			.get(`${url}/_api/search/query?querytext='contentclass:STS_Site'&selectproperties='Title,Path,CreatedOn'&rowlimit=${rowlimit}`, SPHttpClient.configurations.v1)
			.then((res: SPHttpClientResponse): Promise<{ Title: string; Path: string }> => {
				return res.json();
			})
			.then((res: any) => {
				return {
					sites: parseSites(res),
					mySite: getMySite(res),
				};
			})
			.catch(() => {
				return undefined;
			});
	}

	public static async getSitesInfo(spHttpClient: SPHttpClient, sites: SiteInfo[]): Promise<SiteInfo[]> {
		for (let index = 0; index < sites.length; index++) {
			const site = sites[index];
			try {
				const lists = await spClient_Http.getLists(spHttpClient, site.url);
				site.canAccessLists = lists !== undefined && lists?.length > 0;
			} catch {
				/* tslint:disable:no-empty */
			}
		}
		return sortBy(
			sites.filter((s) => s.canAccessLists),
			["name"]
		);
	}

	public static async getLists(spHttpClient: SPHttpClient, siteUrl: string): Promise<ListInfo[] | undefined> {
		const parseLists = (res: any): ListInfo[] => {
			return res.value.map((list: any) => {
				return {
					id: list.Id,
					itemCount: list.ItemCount,
					title: list.Title,
					baseType: list.BaseType,
				};
			});
		};

		return spHttpClient
			.get(`${siteUrl}/_api/web/lists/?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
			.then((res: SPHttpClientResponse): Promise<{ Title: string; Path: string; BaseType: number }> => {
				return res.json();
			})
			.then((res: any) => {
				return parseLists(res);
			})
			.catch(() => {
				return undefined;
			});
	}

	public static async getListContents(spHttpClient: SPHttpClient, siteUrl: string, listId: string, baseType: number): Promise<FileInfo[] | undefined> {
		const getListItems = async (spHttpClient: SPHttpClient, siteUrl: string, listId: string): Promise<FileInfo[] | undefined> => {
			try {
				const results: SPHttpClientResponse = await spHttpClient.get(`${siteUrl}/_api/web/lists/GetById('${listId}')/items`, SPHttpClient.configurations.v1);
				const json = await results.json();
				const contents = json.value;
				return contents.map((res: { Title: string; "@odata.editLink": string; Modified: string }) => {
					return {
						name: res.Title,
						lastModifiedDateTime: moment(res.Modified).format("YYYY.MM.DD HH:mm"),
						webUrl: res["@odata.editLink"],
					} as FileInfo;
				});
			} catch {
				// return undefined;
			}
		};
		const getFiles = async (spHttpClient: SPHttpClient, siteUrl: string, listId: string): Promise<FileInfo[] | undefined> => {
			try {
				const results: SPHttpClientResponse = await spHttpClient.get(`${siteUrl}/_api/web/lists/GetById('${listId}')/files`, SPHttpClient.configurations.v1);
				const json = await results.json();
				const contents = json.value;
				const fileInfo = contents.map((res: { Name: string; Url: string; TimeLastModified: string; Size: number }) => {
					return res.Size > 0 ? ({ name: res.Name, lastModifiedDateTime: moment(res.TimeLastModified).format("YYYY.MM.DD HH:mm"), webUrl: res.Url } as FileInfo) : null;
				});
				return fileInfo.filter((f: FileInfo) => f !== null && f.name.indexOf(".aspx") === -1);
			} catch {
				// return undefined;
			}
		};

		//0:list, 1: library
		return baseType === 0 ? await getListItems(spHttpClient, siteUrl, listId) : await getFiles(spHttpClient, siteUrl, listId);
	}
}
