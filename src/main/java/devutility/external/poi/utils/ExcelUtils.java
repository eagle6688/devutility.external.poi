package devutility.external.poi.utils;

import devutility.internal.io.FileUtils;

public class ExcelUtils {
	/**
	 * Is xls file.
	 * @param fileName: File name
	 * @return boolean
	 */
	public static boolean isXls(String fileName) {
		String extension = FileUtils.getFileExtension(fileName).toLowerCase();
		return ".xls".equals(extension);
	}

	/**
	 * Is xlsx file.
	 * @param fileName: File name.
	 * @return boolean
	 */
	public static boolean isXlsx(String fileName) {
		String extension = FileUtils.getFileExtension(fileName).toLowerCase();
		return ".xlsx".equals(extension);
	}
}