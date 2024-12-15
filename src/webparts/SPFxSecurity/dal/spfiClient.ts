import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import { fileFromAbsolutePath } from "@pnp/sp/files";
import { HttpClient, SPHttpClient } from "@microsoft/sp-http";
import spClient_Http from "./spHttmlClient";
import htmlClientExternalApi from "./htmlClientExternalApi";
import { ApimConfig } from "./types";
export default class spfiClient {
	public static async sendSiteContents(spHttpClient: SPHttpClient, spfiContext: SPFI, httpClient: HttpClient, apimConfig: ApimConfig, siteUrl: string): Promise<void> {
		try {
			const lists = await spClient_Http.getLists(spHttpClient, siteUrl);
			if (lists === undefined) return;

			for (let index = 0; index < lists.length; index++) {
				const list = lists[index];
				if (list.baseType === 1) {
					const fileList = await spClient_Http.getListContents(spHttpClient, siteUrl, list.id, list.baseType);
					if (fileList !== undefined && fileList?.length > 0) {
						console.log(`Sending files from ${list.title}---------------`);

						for (let i = 0; i < fileList.length; i++) {
							const url = fileList[i].webUrl;
							const file = await fileFromAbsolutePath(spfiContext.web, url);
							const fileContent = await file.getBlob();
							//Now send it away
							const res = await htmlClientExternalApi.SendFiles(httpClient, apimConfig, fileList[i].name, fileContent);
							console.log(res);
						}
					}
				}
			}
		} catch (err) {
			console.log(err);
		}
	}
}
