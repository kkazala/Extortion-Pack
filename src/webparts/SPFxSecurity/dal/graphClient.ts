import { sortBy } from "@microsoft/sp-lodash-subset";
import { GraphFI } from "@pnp/graph";
import "@pnp/graph/calendars";
import "@pnp/graph/files";
import "@pnp/graph/groups";
import "@pnp/graph/onenote";
import "@pnp/graph/search";
import "@pnp/graph/sites";
import "@pnp/graph/users";
import "@pnp/graph/lists";
import "@pnp/graph/list-item";
import "@pnp/graph/files";
import "@pnp/graph/mail/messages";
import "@pnp/graph/mail";
import * as moment from "moment";
import { ExoEvent, ExoMail, FileInfo, ListInfo, OneDriveContents, SiteInfo, UserInfo } from "./types";

export default class graphClient {
	private static parseBody(body: { contentType: string; content: string }): string | undefined {
		return body === null
			? undefined
			: body.contentType === "html"
			? body.content
					.replace(/<[^>]*>/g, "")
					.replace(/(\r\n|\n|\r)/gm, "")
					.replace("&nbsp;", "")
					.substring(0, 25)
			: body.content.replace(/(\r\n|\n|\r)/gm, "").substring(0, 25);
	}

	public static async getMe(graph: GraphFI): Promise<UserInfo | undefined> {
		// https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http
		// User.Read
		try {
			const me = await graph.me();
			return {
				businessPhones: me.businessPhones?.join("; ") ?? "",
				displayName: me.displayName ?? "",
				mail: me.mail ?? "",
				officeLocation: me.officeLocation ?? "",
			};
		} catch {
			return undefined;
		}
	}

	public static async getSites(graph: GraphFI): Promise<SiteInfo[] | undefined> {
		// https://learn.microsoft.com/en-us/graph/api/site-search?view=graph-rest-1.0&tabs=http
		// Graph Sites.Read.All
		try {
			const results = await graph.query({
				entityTypes: ["site"],
				query: {
					queryString: "contentclass:STS_Site",
				},
				from: 0,
				size: 100,
				fields: ["displayName", "webUrl", "id"],
			});
			if (results && results.length > 0) {
				const c = results[0].hitsContainers;
				if (c && c.length > 0) {
					const hits = c[0].hits;
					if (hits && hits.length > 0) {
						return hits.map((hit: any) => {
							return {
								name: hit.resource.displayName,
								url: hit.resource.webUrl,
								id: hit.resource.id,
							};
						});
					}
				}
			}
			return [];
		} catch {
			return undefined;
		}
	}

	//returns max 10 sites where contents can be accessed
	public static async getSitesInfo(graph: GraphFI, sites: SiteInfo[]): Promise<SiteInfo[]> {
		// Graph Sites.Read.All or Sites.Selected

		for (let index = 0; index < sites.length; index++) {
			const site = sites[index];
			try {
				const lists = await graph.sites.getById(site.id).lists();
				site.canAccessLists = lists !== null && lists?.length > 0;
			} catch {
				/* tslint:disable:no-empty */
			}
		}
		const _s = sortBy(
			sites.filter((s) => s.canAccessLists),
			["name"]
		);
		return _s.slice(0, 10);
	}
	public static async getLists(graph: GraphFI, siteId: string): Promise<ListInfo[] | undefined> {
		const lists = await graph.sites.getById(siteId).lists();
		const listsInfo = lists
			.filter((l) => l.list?.hidden === false)
			.map((value: any, index: number) => {
				return {
					title: value.displayName,
					id: value.id,
					baseType: value.webUrl.indexOf("/Lists/") === -1 ? 1 : 0,
					itemCount: 0,
				};
			});
		for (let index = 0; index < listsInfo.length; index++) {
			const list = listsInfo[index];
			const items = await graph.sites.getById(siteId).lists.getById(list.id).items();
			list.itemCount = items.length;
		}

		return listsInfo;
	}
	public static async getListContents(graph: GraphFI, siteId: string, listId: string, baseType: number): Promise<FileInfo[] | undefined> {
		const getListItems = async (graph: GraphFI, siteId: string, listId: string): Promise<FileInfo[] | undefined> => {
			const items = await graph.sites.getById(siteId).lists.getById(listId).items.expand("fields")();
			return items.map((value: any) => {
				return {
					name: value.fields.LinkTitle,
					lastModifiedDateTime: moment(value.TimeLastModified).format("YYYY.MM.DD HH:mm"),
					webUrl: value.webUrl,
				} as FileInfo;
			});
		};
		const getFiles = async (graph: GraphFI, siteId: string, listId: string): Promise<FileInfo[] | undefined> => {
			const files = await graph.sites.getById(siteId).lists.getById(listId).items.expand("fields")();
			return files
				.filter((f) => f.contentType?.name === "Document")
				.map((value: any) => {
					return {
						name: value.fields.FileLeafRef,
						lastModifiedDateTime: moment(value.TimeLastModified).format("YYYY.MM.DD HH:mm"),
						webUrl: value.webUrl,
					} as FileInfo;
				});
		};
		//0:list, 1: library
		return baseType === 0 ? await getListItems(graph, siteId, listId) : await getFiles(graph, siteId, listId);
	}
	public static async getFilesSharedWithMe(graph: GraphFI): Promise<FileInfo[] | undefined> {
		//Files.Read
		try {
			const files = await graph.me.drive.sharedWithMe();

			return Object(files)
				.slice(0, 10)
				.map((file: any) => {
					return {
						name: file.name,
						lastModifiedDateTime: new Date(file.lastModifiedDateTime).toLocaleDateString(),
						webUrl: file.webUrl,
					};
				});
		} catch (error) {
			return undefined;
			// console.clear();
		}
	}

	public static async getOneDriveRecentFiles(graph: GraphFI): Promise<FileInfo[] | undefined> {
		//Files.Read
		try {
			const files = await graph.me.drive.recent(); //name, lastModifiedDateTime,webUrl

			return files.slice(0, 10).map((file: any) => {
				return {
					name: file.name,
					lastModifiedDateTime: new Date(file.lastModifiedDateTime).toLocaleDateString(),
					webUrl: file.webUrl,
				};
			});
		} catch (error) {
			return undefined;
			// console.clear();
		}
	}
	public static async getOneDriveContents(graph: GraphFI): Promise<OneDriveContents[] | undefined> {
		//Files.Read
		try {
			const rootChildren = await graph.me.drive.root.children(); //name, webUrl, folder.ChildCount

			return rootChildren.slice(0, 10).map((folder: any) => {
				return {
					name: folder.name,
					webUrl: folder.webUrl,
					childCount: folder.folder ? folder.folder.childCount : 0,
				};
			});
		} catch (error) {
			return undefined;
			// console.clear();
		}
	}
	public static async getEvents(graph: GraphFI): Promise<ExoEvent[] | undefined> {
		//	Calendars.ReadBasic, Calendars.Read
		try {
			const monday = moment().weekday(1).format("YYYY-MM-DD");
			const friday = moment().weekday(5).format("YYYY-MM-DD");
			const eventsThisWeek = await graph.me.calendarView(monday, friday).select("subject", "start", "end", "isAllDay", "body").top(10)();

			return eventsThisWeek.map((event: any) => {
				const isSameDay = moment(event.start.dateTime).isSame(moment(event.end.dateTime));
				const isAllDay = event.isAllDay;
				const startEnd = isSameDay
					? isAllDay
						? moment(event.start.dateTime).format("YYYY.MM.DD")
						: `${moment(event.start.dateTime).format("YYYY.MM.DD")} ${moment(event.start.dateTime).format("HH:mm")}-${moment(event.end.dateTime).format("HH:mm")}`
					: isAllDay
					? `${moment(event.start.dateTime).format("YYYY.MM.DD")} - ${moment(event.end.dateTime).format("YYYY.MM.DD")}`
					: `${moment(event.start.dateTime).format("YYYY.MM.DD HH:mm")} - ${moment(event.end.dateTime).format("YYYY.MM.DD HH:mm")}`;

				return {
					subject: event.subject,
					isAllDay: event.isAllDay,
					startEnd,
					body: graphClient.parseBody(event.body),
				};
			});
		} catch (error) {
			console.clear();
			return undefined;
		}
	}
	public static async getAllCalendars(graph: GraphFI): Promise<string[] | undefined> {
		//Calendars.ReadBasic
		try {
			const calendars = await graph.me.calendars();

			return calendars.map((calendar: any) => calendar.name);
		} catch (error) {
			console.clear();
			return undefined;
		}
	}
	public static async getEmails(graph: GraphFI): Promise<ExoMail[] | undefined> {
		// Mail.ReadBasic;
		// User.Read
		try {
			const messages = await graph.me.messages();

			return messages.slice(0, 10).map((mail: any) => {
				return {
					subject: mail.subject,
					from: mail.sender.emailAddress.name,
					body: graphClient.parseBody(mail.body),
				};
			});
		} catch (error) {
			console.clear();
			return undefined;
		}
	}

	public static async getUsers(graph: GraphFI): Promise<UserInfo[] | undefined> {
		//User.ReadBasic.All, User.Read.All,
		try {
			const allUsers = await graph.users();

			return allUsers.slice(0, 10).map((user: any) => {
				return {
					businessPhones: user.businessPhones.join("; "),
					displayName: user.displayName,
					mail: user.mail,
					officeLocation: user.officeLocation,
				};
			});
		} catch (error) {
			console.clear();
			return undefined;
		}
	}
}
