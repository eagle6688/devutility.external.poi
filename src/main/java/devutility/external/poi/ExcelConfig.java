package devutility.external.poi;

import devutility.external.poi.models.ExcelProperties;

public class ExcelConfig {
	/**
	 * Max sheets count.
	 */
	public final static int MAXSHEETSCOUNT = 255;

	/**
	 * Max rows count for .xls files.
	 */
	public final static int MAXROWSCOUNTPRE07 = 65536;

	/**
	 * Max columns count for .xls files.
	 */
	public final static int MAXCOLUMNSCOUNTPRE07 = 256;

	/*
	 * Max rows count for .xlsx files.
	 */
	public final static int MAXROWSCOUNTAFTER07 = 1048576;

	/**
	 * Max columns count for .xlsx files.
	 */
	public final static int MAXCOLUMNSCOUNTAFTER07 = 16384;

	public static ExcelProperties ExcelPropertiesAfter07 = new ExcelProperties() {
	};
}