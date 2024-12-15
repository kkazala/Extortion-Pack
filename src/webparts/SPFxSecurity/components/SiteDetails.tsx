import { Button, FlatTree, Link, TableCell, TableCellActions, TableRow, Tooltip } from "@fluentui/react-components";
import { EmojiAngryFilled, EyeFilled, EyeOffFilled } from "@fluentui/react-icons";
import * as React from "react";
import { SPHttpClient } from "@microsoft/sp-http";
import SiteListDetails from "./SiteListDetails";
import { SiteInfo, ListInfo, FileInfo } from "../dal/types";
import useStyles from "./styles";

export type SiteDetailsRowProps = {
	value: SiteInfo;
	index: number;
	isApimEnabled?: boolean;
	getLists(): Promise<ListInfo[] | undefined>;
	showContents(listId: string, baseType: number): Promise<FileInfo[] | undefined>;
	sendContents?(): Promise<void>;

	spHttpClient?: SPHttpClient;
};

const SiteDetails = (props: SiteDetailsRowProps): JSX.Element => {
	const styles = useStyles();
	const { index, value, isApimEnabled } = props;
	const [spListsInfo, setspListInfo] = React.useState<ListInfo[] | undefined>();
	const [showDetails, setShowDetails] = React.useState<boolean>(false);

	const tooltip: string[] = [
		"Click me. #YOLO",
		"Click me. Don't worry about it...",
		"Click me. No risk no fun!",
		"Click me. What's the worst that can happen?",
		"Click me. Stop thinking about it...",
		"Hackers won't ask you to click a button...",
	];
	const getRandomTooltip = (): string => {
		const min = 0;
		const max = tooltip.length - 1;
		const rand = Math.floor(Math.random() * (max - min + 1)) + min;
		return tooltip[rand];
	};
	const showContents = async (): Promise<void> => {
		if (spListsInfo === undefined) {
			const _lists = await props.getLists();
			setspListInfo(_lists);
		}
		setShowDetails(!showDetails);
	};

	return (
		<>
			<TableRow key={`s${index}`}>
				<TableCell>
					<Tooltip
						withArrow
						content={{ children: value.url, className: styles.tooltip }}
						relationship="description"
					>
						<Link
							href={value.url}
							target="_blank"
						>
							{value.name}
						</Link>
					</Tooltip>
					{props.sendContents !== undefined && isApimEnabled && (
						<TableCellActions>
							<Tooltip
								withArrow
								content={{ children: getRandomTooltip(), className: styles.tooltip }}
								relationship="description"
							>
								<Button
									onClick={props.sendContents}
									icon={<EmojiAngryFilled className={styles.iconSend} />}
									appearance="subtle"
									aria-label={"Leak list contents"}
								/>
							</Tooltip>
						</TableCellActions>
					)}
				</TableCell>
				<TableCell>
					<Tooltip
						withArrow
						content={{ children: "Show site contents", className: styles.tooltip }}
						relationship="label"
					>
						<Button
							onClick={showContents}
							icon={showDetails ? <EyeOffFilled /> : <EyeFilled />}
							appearance="subtle"
							aria-label={showDetails ? "Hide contents" : "Show contents"}
						/>
					</Tooltip>
				</TableCell>
			</TableRow>
			<TableRow key={`details${index}`}>
				{showDetails && spListsInfo !== undefined && spListsInfo.length > 0 && (
					<TableCell colSpan={2}>
						<FlatTree>
							{spListsInfo.map((list: ListInfo, index: number) => {
								return (
									<SiteListDetails
										key={`s${index}`}
										index={index}
										listInfo={list}
										siteUrl={value.url}
										showContents={() => props.showContents(list.id, list.baseType)}
									/>
								);
							})}
						</FlatTree>
					</TableCell>
				)}
			</TableRow>
		</>
	);
};

export default SiteDetails;
