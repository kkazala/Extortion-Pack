import { FlatTreeItem, TreeItemLayout, CounterBadge, TreeItemOpenChangeEvent, TreeItemOpenChangeData, Spinner } from "@fluentui/react-components";
import * as React from "react";
import { ListInfo, FileInfo } from "../dal/types";

export type SiteListDetailsProps = {
	siteUrl: string;
	listInfo: ListInfo;
	index: number;
	showContents(siteRef: string, listId: string, baseType: number): Promise<FileInfo[] | undefined>;
};

const SiteListDetails = (props: SiteListDetailsProps): JSX.Element => {
	const { siteUrl, listInfo, index } = props;
	const [open, setOpen] = React.useState(false);
	const [isLoading, setIsLoading] = React.useState(false);
	const [spListItems, setListItems] = React.useState<FileInfo[] | undefined>();
	const firstItemRef = React.useRef<HTMLDivElement>(null);

	const showContents = async (listId: string, baseType: number, isOpen: boolean): Promise<void> => {
		if (isOpen) {
			setIsLoading(true);
			const content = await props.showContents(siteUrl, listId, baseType);
			setListItems(content);

			setIsLoading(false);
		}
		setOpen(isOpen);
	};

	return (
		<>
			<FlatTreeItem
				itemType={listInfo.itemCount > 0 ? "branch" : "leaf"}
				aria-level={1}
				aria-posinset={index + 1}
				aria-setsize={listInfo.itemCount}
				value={index}
				open={open}
				onOpenChange={(e: TreeItemOpenChangeEvent, data: TreeItemOpenChangeData) => {
					// eslint-disable-next-line @typescript-eslint/no-floating-promises
					showContents(listInfo.id, listInfo.baseType, data.open);
				}}
			>
				<TreeItemLayout
					aside={
						<CounterBadge
							count={listInfo.itemCount}
							appearance="filled"
							size="small"
							color="informative"
							showZero={true}
						/>
					}
					expandIcon={isLoading ? <Spinner size="tiny" /> : undefined}
				>
					{listInfo.title}
				</TreeItemLayout>
			</FlatTreeItem>
			{open &&
				spListItems?.map((value: FileInfo, idx: number) => (
					<FlatTreeItem
						key={`con${idx}`}
						ref={index === 0 ? firstItemRef : null}
						parentValue={index}
						value={idx}
						aria-level={2}
						aria-setsize={spListItems.length}
						aria-posinset={idx + 1}
						itemType="leaf"
					>
						<TreeItemLayout>{value.name}</TreeItemLayout>
					</FlatTreeItem>
				))}
		</>
	);
};

export default SiteListDetails;
