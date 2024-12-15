import {
	Button,
	Menu,
	MenuItem,
	MenuItemProps,
	MenuList,
	MenuPopover,
	MenuTrigger,
	Overflow,
	OverflowItem,
	Tab,
	TabList,
	makeStyles,
	tokens,
	useIsOverflowItemVisible,
	useOverflowMenu,
} from "@fluentui/react-components";
import { MoreHorizontalFilled } from "@fluentui/react-icons";
import * as React from "react";

type TabsProps = {
	id: string;
	name: string;
	visible: boolean;
};

const useStyles = makeStyles({
	tabs: {
		paddingBottom: "20px",
	},
});

const tabs: TabsProps[] = [
	{
		id: "about",
		name: "Overview",
		visible: true,
	},
	{
		id: "restAPI",
		name: "SharePoint REST API",
		visible: false,
	},
	{
		id: "graphAPISites",
		name: "Graph API: SharePoint",
		visible: false,
	},
	{
		id: "graphAPIEXO",
		name: "Graph API: Mail & Calendar",
		visible: false,
	},
	{
		id: "graphAPIUsers",
		name: "Graph API: Users",
		visible: false,
	},
];

//#region OverflowMenuItem
type OverflowMenuItemProps = {
	tab: TabsProps;
	onClick: MenuItemProps["onClick"];
};

/**
 * A menu item for an overflow menu that only displays when the tab is not visible
 */
const OverflowMenuItem = (props: OverflowMenuItemProps): JSX.Element | null => {
	const { tab, onClick } = props;
	const isVisible = useIsOverflowItemVisible(tab.id);

	if (isVisible) {
		return null;
	}

	return (
		<MenuItem
			key={tab.id}
			onClick={onClick}
		>
			<div>{tab.name}</div>
		</MenuItem>
	);
};
//#endregion

//#region  OverflowMenu

const useOverflowMenuStyles = makeStyles({
	menu: {
		backgroundColor: tokens.colorNeutralBackground1,
	},
	menuButton: {
		alignSelf: "center",
	},
});

type OverflowMenuProps = {
	onTabSelect?: (tabId: string) => void;
	tabs: TabsProps[];
};

/**
 * A menu for selecting tabs that have overflowed and are not visible.
 */
const OverflowMenu = (props: OverflowMenuProps): JSX.Element | null => {
	const { onTabSelect, tabs } = props;
	const { ref, isOverflowing, overflowCount } = useOverflowMenu<HTMLButtonElement>();

	const styles = useOverflowMenuStyles();

	const onItemClick = (tabId: string): void => {
		onTabSelect?.(tabId);
	};

	if (!isOverflowing) {
		return null;
	}

	return (
		<Menu hasIcons>
			<MenuTrigger disableButtonEnhancement>
				<Button
					appearance="transparent"
					className={styles.menuButton}
					ref={ref}
					icon={<MoreHorizontalFilled />}
					aria-label={`${overflowCount} more tabs`}
					role="tab"
				/>
			</MenuTrigger>
			<MenuPopover>
				<MenuList className={styles.menu}>
					{tabs.map((tab) => (
						<OverflowMenuItem
							key={tab.id}
							tab={tab}
							onClick={() => onItemClick(tab.id)}
						/>
					))}
				</MenuList>
			</MenuPopover>
		</Menu>
	);
};
//#endregion

export type TabListMenuProps = {
	isLoading: boolean;
	onTabSelected: (tabId: string) => void;
};

const TabListMenu = (props: TabListMenuProps): JSX.Element => {
	const { onTabSelected } = props;
	const [selectedTabId, setSelectedTabId] = React.useState<string>("about");
	const [refreshKey, setRefreshKey] = React.useState<number>(Date.now());
	const styles = useStyles();

	//show tabs if operations completed
	React.useEffect(() => {
		if (props.isLoading === false) {
			tabs[1].visible = true; //"SharePoint REST API"
			tabs[2].visible = true; //"Microsoft Graph: SharePoint"
			tabs[3].visible = true; //"Microsoft Graph: Mail & Calendar"
			tabs[4].visible = true; //"Microsoft Graph: Users"

			setRefreshKey(Date.now());
		}
	}, [props.isLoading]);

	const onTabSelect = (tabId: string): void => {
		setSelectedTabId(tabId);
		onTabSelected(tabId);
	};
	return (
		<div>
			<Overflow minimumVisible={5}>
				<TabList
					size="large"
					key={refreshKey}
					selectedValue={selectedTabId}
					onTabSelect={(_, d) => onTabSelect(d.value as string)}
					className={styles.tabs}
				>
					{tabs
						.filter((t) => t.visible === true)
						.map((tab) => {
							return (
								<OverflowItem
									key={tab.id}
									id={tab.id}
									priority={tab.id === selectedTabId ? 2 : 1}
								>
									<Tab value={tab.id}>{tab.name}</Tab>
								</OverflowItem>
							);
						})}
					<OverflowMenu
						onTabSelect={onTabSelect}
						tabs={tabs}
					/>
				</TabList>
			</Overflow>
		</div>
	);
};

export default TabListMenu;
