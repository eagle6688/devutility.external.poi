package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

public class SheetUtils {
	/**
	 * Create a Sheet object with unsafe name.
	 * @param workbook: Workbook object.
	 * @param name: Sheet name.
	 * @return Sheet
	 */
	public static Sheet createSheet(Workbook workbook, String name) {
		String safeName = WorkbookUtil.createSafeSheetName(name);
		return workbook.createSheet(safeName);
	}
}