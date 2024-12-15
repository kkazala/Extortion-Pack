import * as React from "react";

import { Tooltip } from "@fluentui/react-components";
import { ShieldErrorFilled, ShieldTaskFilled } from "@fluentui/react-icons";
import useStyles from "./styles";

export type ShieldProps = {
	className: string;
	text?: string;
};

const ShieldError = (props: ShieldProps): JSX.Element => {
	const styles = useStyles();
	return (
		<Tooltip
			withArrow
			content={{ children: props.text ?? "This information may be accessed by all SPFx solutions", className: styles.tooltip }}
			relationship="description"
		>
			<ShieldErrorFilled className={props.className} />
		</Tooltip>
	);
};
const ShieldOK = (props: ShieldProps): JSX.Element => {
	const styles = useStyles();
	return (
		<Tooltip
			withArrow
			content={{ children: props.text ?? "This information is inaccessible to SPFx solutions", className: styles.tooltip }}
			relationship="description"
		>
			<ShieldTaskFilled className={props.className} />
		</Tooltip>
	);
};

export { ShieldError, ShieldOK };
