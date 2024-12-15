import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApimConfig } from "../dal/types";

export type SPFxSecurityProps = {
	context: WebPartContext;
	apimConfig?: ApimConfig;
};
