import { tokens, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
	narrowColumn: {
		width: "200px",
	},
	tableHeader: {
		backgroundColor: tokens.colorNeutralBackground1Hover,
	},
	moreInfo: {
		backgroundColor: tokens.colorNeutralBackground1Hover,
		padding: "20px",
	},
	warning: {
		color: tokens.colorPaletteRedForeground2,
	},
	addPadding: {
		padding: "20px 0",
	},
	addMargin: {
		margin: "20px 0",
	},
	icon: { fontSize: "20px", verticalAlign: "bottom" },
	btnCopy: { height: "32px" },
	iconCopy: { color: tokens.colorBrandBackground },
	tooltip: {
		//currently, the tokens.*** background color, and font family and size are not applied correctly
		backgroundColor: "#fafafa",
		fontFamily: "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif",
		fontSize: "14px",
		color: tokens.colorNeutralForegroundInverted,
		padding: "8px",
	},
	textCode: {
		backgroundColor: "#fafafa",
		fontFamily: "monospace",
	},
});

export default useStyles;
