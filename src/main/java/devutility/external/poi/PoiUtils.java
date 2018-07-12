package devutility.external.poi;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

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
	public static <T> List<T> read(String filePath, String sheet, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws Exception {
		List<T> list = new LinkedList<>();
		Workbook workbook = WorkbookUtils.load(filePath);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		list = SheetUtils.toList(sheetObj, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}

	/**
	 * Read data from InputStream and convert data to List (Type T).
	 * @param inputStream: InputStream for excel file.
	 * @param sheet: Sheet name in excel file.
	 * @param fieldColumnMap: The map between type T field and excel column.
	 * @param clazz: Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static <T> List<T> read(InputStream inputStream, String sheet, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws Exception {
		List<T> list = new LinkedList<>();
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		list = SheetUtils.toList(sheetObj, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}

	public static <T> OutputStream save(InputStream templateInputStream, String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list) {
		return null;
	}

	public static <T> OutputStream save(String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list) {
		return null;
	}
}