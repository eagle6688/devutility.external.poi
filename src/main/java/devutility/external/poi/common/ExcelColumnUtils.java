package devutility.external.poi.common;

import java.lang.reflect.Field;

/**
 * 
 * ExcelColumnUtils
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 19:56:38
 */
public class ExcelColumnUtils {
	/**
	 * Return index number of excel column, -1 means didn't set.
	 * @param field Field object.
	 * @return int
	 */
	public static int getIndex(Field field) {
		ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);

		if (excelColumn == null) {
			return -1;
		}

		return excelColumn.index();
	}
}