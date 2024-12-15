import { tokens, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
	iconSend: {
		color: tokens.colorBrandBackground,
		fontSize: "32px",
		":hover": {
			color: tokens.colorStatusDangerBackground3,
		},
	},
	tooltip: {
		//currently, the tokens.*** background color, and font family and size are not applied correctly
		backgroundColor: "#fafafa",
		fontFamily: "'Segoe UI', 'Segoe UI Web (West European)', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif",
		fontSize: "14px",
		color: tokens.colorNeutralForegroundInverted,
		padding: "8px",
	},

	icon24OK: { fontSize: "24px", color: tokens.colorPaletteGreenForeground1 },
	icon24Warn: { fontSize: "24px", color: tokens.colorPaletteDarkOrangeForeground1 },
	icon24: { fontSize: "24px" },
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
});

export default useStyles;
