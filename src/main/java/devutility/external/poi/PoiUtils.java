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
import devutility.external.poi.models.RowStyle;
import devutility.external.poi.utils.SheetUtils;
import devutility.external.poi.utils.WorkbookUtils;

public class PoiUtils {
	/**
	 * Read data from file and convert data to List (Type T).
	 * @param filePath: Excel file path.
	 * @param sheetName: Sheet name in Excel file.
	 * @param fieldColumnMap: The map between type T field and Excel column.
	 * @param clazz: Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static <T> List<T> read(String filePath, String sheetName, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws Exception {
		Workbook workbook = WorkbookUtils.load(filePath);
		Sheet sheet = SheetUtils.get(workbook, sheetName);
		List<T> list = SheetUtils.toList(sheet, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}

	/**
	 * Read data from InputStream and convert data to List (Type T).
	 * @param inputStream: InputStream for Excel file.
	 * @param sheetName: Sheet name in Excel file.
	 * @param fieldColumnMap: The map between type T field and Excel column.
	 * @param clazz: Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static <T> List<T> read(InputStream inputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = SheetUtils.get(workbook, sheetName);
		List<T> list = SheetUtils.toList(sheet, fieldColumnMap, clazz);
		workbook.close();
		return list;
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream
	 * object.
	 * @param templateInputStream: InputStream for template.
	 * @param sheetName: Sheet name in template and will in outputStream.
	 * @param fieldColumnMap
	 * @param list: List data.
	 * @param rowStyle: Style for row.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, RowStyle rowStyle, OutputStream outputStream) throws Exception {
		Workbook workbook = WorkbookUtils.save(templateInputStream, sheetName, fieldColumnMap, list, rowStyle);
		workbook.write(outputStream);
		workbook.close();
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream
	 * object.
	 * @param templateInputStream: InputStream for template.
	 * @param sheetName: Sheet name in template and will in outputStream.
	 * @param fieldColumnMap
	 * @param list: List data.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, OutputStream outputStream) throws Exception {
		save(templateInputStream, sheetName, fieldColumnMap, list, null, outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path,
	 * save list data into it.
	 * @param templateInputStream: InputStream for template.
	 * @param sheetName: Sheet name in template and will in output Excel file.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param rowStyle: Style for row.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, RowStyle rowStyle, String filePath) throws Exception {
		FileOutputStream outputStream = createFileOutputStream(filePath);
		save(templateInputStream, sheetName, fieldColumnMap, list, rowStyle, outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path,
	 * save list data into it.
	 * @param templateInputStream: InputStream for template.
	 * @param sheetName: Sheet name in template and will in output Excel file.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		save(templateInputStream, sheetName, fieldColumnMap, list, null, filePath);
	}

	/**
	 * Save the list data into specific OutputStream.
	 * @param excelType: ExcelType object.
	 * @param sheetName: Sheet name in outputStream.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, OutputStream outputStream) throws Exception {
		Workbook workbook = WorkbookUtils.save(excelType, sheetName, fieldColumnMap, list);
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
		String sheetName = SheetUtils.getName(1);
		save(excelType, sheetName, fieldColumnMap, list, outputStream);
	}

	/**
	 * Create a new Excel file with specific excelType, sheet name and file path,
	 * save list data into it.
	 * @param excelType: Excel type.
	 * @param sheetName: Sheet name.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		FileOutputStream outputStream = createFileOutputStream(filePath);
		save(excelType, sheetName, fieldColumnMap, list, outputStream);
	}

	/**
	 * Create a new Excel file with specific excelType and file path, save list data
	 * into it.
	 * @param excelType: Excel type.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		String sheetName = SheetUtils.getName(1);
		save(excelType, sheetName, fieldColumnMap, list, filePath);
	}

	/**
	 * Create a new Excel 2007 file with specific excelType, sheet name and file
	 * path, save list data into it.
	 * @param sheetName: Sheet name in Excel file.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		save(ExcelType.Excel2007, sheetName, fieldColumnMap, list, filePath);
	}

	/**
	 * Create a new Excel 2007 file with specific excelType and file path, save list
	 * data into it.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(FieldColumnMap<T> fieldColumnMap, List<T> list, String filePath) throws Exception {
		save(ExcelType.Excel2007, fieldColumnMap, list, filePath);
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream
	 * object.
	 * @param templateWorkbook: Workbook for template.
	 * @param sheetName: Sheet name in template and will in outputStream.
	 * @param fieldColumnMap
	 * @param list: List data.
	 * @param rowStyle: Style for row.
	 * @param outputStream: OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(Workbook templateWorkbook, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, RowStyle rowStyle, OutputStream outputStream) throws Exception {
		Workbook workbook = WorkbookUtils.save(templateWorkbook, sheetName, fieldColumnMap, list, rowStyle);
		workbook.write(outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path,
	 * save list data into it.
	 * @param templateWorkbook: Workbook for template.
	 * @param sheetName: Sheet name in template and will in output Excel file.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param rowStyle: Style for row.
	 * @param filePath: File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(Workbook templateWorkbook, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, RowStyle rowStyle, String filePath) throws Exception {
		FileOutputStream outputStream = createFileOutputStream(filePath);
		save(templateWorkbook, sheetName, fieldColumnMap, list, rowStyle, outputStream);
	}

	/**
	 * Create an FileOutputStream instance.
	 * @param filePath: File path.
	 * @return FileOutputStream
	 * @throws FileNotFoundException: Throw while file not found.
	 */
	private static FileOutputStream createFileOutputStream(String filePath) throws FileNotFoundException {
		File file = new File(filePath);
		return new FileOutputStream(file);
	}
}