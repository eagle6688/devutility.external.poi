package devutility.external.poi;

import devutility.external.poi.models.ExcelProperties;

public class ExcelConfig {
	/**
	 * ExcelProperties for version before 2007.
	 */
	public static ExcelProperties EXCELPROPERTIESPRE07 = new ExcelProperties(255, 65536, 256);

	/**
	 * ExcelProperties for version after 2007.
	 */
	public static ExcelProperties EXCELPROPERTIESAFTER07 = new ExcelProperties(255, 1048576, 16384);
}