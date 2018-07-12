package devutility.external.poi;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.common.ExcelType;
import devutility.external.poi.models.FieldColumnMap;
import devutility.external.poi.utils.SheetUtils;
import devutility.external.poi.utils.WorkbookUtils;
import devutility.internal.io.FileUtils;

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
		Workbook workbook = WorkbookUtils.load(filePath);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		List<T> list = SheetUtils.toList(sheetObj, fieldColumnMap, clazz);
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
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		List<T> list = SheetUtils.toList(sheetObj, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}

	/**
	 * Append list data into template excel file and save them into OutputStream
	 * object.
	 * @param templateInputStream: InputStream for template.
	 * @param sheet: Sheet name need append.
	 * @param fieldColumnMap
	 * @param list: List data.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list, OutputStream outputStream) throws Exception {
		Workbook workbook = WorkbookFactory.create(templateInputStream);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		SheetUtils.append(sheetObj, fieldColumnMap, list);
		workbook.write(outputStream);
		workbook.close();
	}

	/**
	 * Save the list data into specific OutputStream.
	 * @param excelType: ExcelType object.
	 * @param sheet: Sheet name.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list, OutputStream outputStream) throws Exception {
		Workbook workbook = WorkbookUtils.create(excelType);
		Sheet sheetObj = SheetUtils.get(workbook, sheet);
		SheetUtils.append(sheetObj, fieldColumnMap, list);
		workbook.write(outputStream);
		workbook.close();
	}

	/**
	 * Save the list data into specific OutputStream.
	 * @param excelType: ExcelType object.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, FieldColumnMap<T> fieldColumnMap, List<T> list, OutputStream outputStream) throws Exception {
		String sheet = SheetUtils.getName(1);
		save(excelType, sheet, fieldColumnMap, list, outputStream);
	}

	/**
	 * Create a new excel file with specific template, sheet name and file path,
	 * save list data into it.
	 * @param templateInputStream: InputStream for template.
	 * @param sheet: Sheet name.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new excel file.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		FileOutputStream outputStream = createFileOutputStream(filePath);
		save(templateInputStream, sheet, fieldColumnMap, list, outputStream);
	}

	/**
	 * Create a new excel file with specific excelType, sheet name and file path,
	 * save list data into it.
	 * @param excelType: Excel type.
	 * @param sheet: Sheet name.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new excel file.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, String sheet, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		FileOutputStream outputStream = createFileOutputStream(filePath);
		save(excelType, sheet, fieldColumnMap, list, outputStream);
	}

	/**
	 * Create a new excel file with specific excelType and file path, save list data
	 * into it.
	 * @param excelType: Excel type.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new excel file.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		String sheet = SheetUtils.getName(1);
		save(excelType, sheet, fieldColumnMap, list, filePath);
	}

	/**
	 * Create an FileOutputStream instance.
	 * @param filePath: File path.
	 * @return FileOutputStream
	 * @throws FileNotFoundException: Throw while file not found.
	 */
	private static FileOutputStream createFileOutputStream(String filePath) throws FileNotFoundException {
		File file = FileUtils.create(filePath);
		return new FileOutputStream(file);
	}
}