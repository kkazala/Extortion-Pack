import { HttpClient, IHttpClientOptions } from "@microsoft/sp-http";
import { ApimConfig } from "./types";

export default class htmlClientExternalApi {
	public static async Send(httpClient: HttpClient, testVal: string): Promise<string | null> {
		const postURL = "https://httpbin.org/post";
		const requestHeaders: Headers = new Headers();
		const httpClientOptions: IHttpClientOptions = {
			body: testVal,
			headers: requestHeaders,
			mode: "cors",
			method: "post",
		};

		const res = await httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions);
		const response = await res.json();
		return response?.data ?? null;
	}

	public static async SendFiles(httpClient: HttpClient, apimConfig: ApimConfig, fileName: string, fileContent: Blob): Promise<string | void> {
		const postURL = apimConfig.endpoint;
		// (tenant sws2)
		const requestHeaders: Headers = new Headers();
		requestHeaders.append("ocp-apim-subscription-key", apimConfig.key);
		requestHeaders.append("Content-Type", "multipart/form-data");
		requestHeaders.append("Origin", "https://spfx.com");

		const formData = new FormData();
		formData.set("file", fileContent, fileName);

		const httpClientOptions: IHttpClientOptions = {
			headers: requestHeaders,
			mode: "cors",
			method: "post",
			body: formData,
		};

		try {
			const res = await httpClient.post(postURL, HttpClient.configurations.v1, httpClientOptions);

			return res.ok ? `Sucesfully sent ${fileName}` : `Something went wrong when sending ${fileName}`;
		} catch (err) {
			console.log(err);
		}
	}
}
