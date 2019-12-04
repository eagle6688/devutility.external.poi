package devutility.external.poi.common;

import devutility.external.poi.model.ExcelProperties;

/**
 * 
 * ExcelConfig
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 18:52:17
 */
public class ExcelConfig {
	/**
	 * ExcelProperties for version before 2007.
	 */
	public static ExcelProperties EXCELPROPERTIESPRE07 = new ExcelProperties(255, 65536, 256);

	/**
	 * ExcelProperties for version after 2007.
	 */
	public static ExcelProperties EXCELPROPERTIESAFTER07 = new ExcelProperties(255, 1048576, 16384);

	/**
	 * Sheet name format.
	 */
	public static final String SHEETNAMEFORMAT = "Sheet%d";
}