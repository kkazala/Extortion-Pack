import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApimConfig } from "../dal/types";

export type GetTokensProps = {
	context: WebPartContext;
	apimConfig?: ApimConfig;
};
