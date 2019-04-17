package devutility.external.poi.utils;

import devutility.external.poi.common.ExcelType;
import devutility.internal.io.FileUtils;

public class ExcelUtils {
	/**
	 * Is xls file.
	 * @param fileName File name
	 * @return boolean
	 */
	public static boolean isXls(String fileName) {
		String extension = FileUtils.getExtension(fileName).toLowerCase();
		return ".xls".equals(extension);
	}

	/**
	 * Is xlsx file.
	 * @param fileName File name.
	 * @return boolean
	 */
	public static boolean isXlsx(String fileName) {
		String extension = FileUtils.getExtension(fileName).toLowerCase();
		return ".xlsx".equals(extension);
	}

	/**
	 * Get ExcelType by file name.
	 * @param fileName File name.
	 * @return ExcelType
	 */
	public static ExcelType getExcelType(String fileName) {
		if (isXlsx(fileName)) {
			return ExcelType.Excel2007;
		}

		return ExcelType.Excel2003;
	}
}