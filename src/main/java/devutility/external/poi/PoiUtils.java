package devutility.external.poi;

import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import devutility.external.poi.models.FieldColumnMap;
import devutility.external.poi.utils.SheetUtils;
import devutility.external.poi.utils.WorkbookUtils;

public class PoiUtils {
	/**
	 * Read data from file and convert data to List (Type T).
	 * @param filePath: Excel file path.
	 * @param sheet: Sheet name in excel file.
	 * @param fieldColumnMap: The map between type T field and excel column.
	 * @param clazz: Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static List<T> read(String filePath, String sheet, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws Exception {
		List<T> list = new LinkedList<>();
		Workbook workbook = WorkbookUtils.load(filePath);
		Sheet sheetObj = SheetUtils.getSheet(workbook, sheet);
		list = SheetUtils.toList(sheetObj, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}
}