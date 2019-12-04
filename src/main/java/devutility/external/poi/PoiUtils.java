package devutility.external.poi;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.common.ExcelType;
import devutility.external.poi.model.ColumnFieldMap;
import devutility.external.poi.model.RowStyle;
import devutility.external.poi.utils.ExcelUtils;
import devutility.external.poi.utils.SheetUtils;
import devutility.external.poi.utils.WorkbookUtils;
import devutility.internal.util.CollectionUtils;

/**
 * 
 * PoiUtils
 * 
 * @author: Aldwin Su
 * @version: 2019-09-12 23:02:06
 */
public class PoiUtils {
	/**
	 * Read data from file and convert data to List (Type T).
	 * @param filePath Excel file path.
	 * @param sheetName Sheet name in Excel file.
	 * @param clazz Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static <T> List<T> read(String filePath, String sheetName, Class<T> clazz) throws Exception {
		Workbook workbook = WorkbookUtils.load(filePath);
		Sheet sheet = SheetUtils.get(workbook, sheetName);
		List<T> list = SheetUtils.toList(sheet, new ColumnFieldMap(clazz), clazz);
		workbook.close();
		return list;
	}

	/**
	 * Read data from InputStream and convert data to List (Type T).
	 * @param inputStream InputStream for Excel file.
	 * @param sheetName Sheet name in Excel file.
	 * @param clazz: Class object.
	 * @return {@code List<T>}
	 * @throws Exception
	 */
	public static <T> List<T> read(InputStream inputStream, String sheetName, Class<T> clazz) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = SheetUtils.get(workbook, sheetName);
		List<T> list = SheetUtils.toList(sheet, new ColumnFieldMap(clazz), clazz);
		workbook.close();
		return list;
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream object.
	 * @param templateInputStream InputStream for template.
	 * @param sheetName Sheet name in template and will in outputStream.
	 * @param list {@code List<T>} object.
	 * @param rowStyle Style for row.
	 * @param outputStream OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, List<T> list, RowStyle rowStyle, OutputStream outputStream) throws Exception {
		if (CollectionUtils.isNullOrEmpty(list)) {
			return;
		}

		Class<?> clazz = list.get(0).getClass();
		ColumnFieldMap columnFieldMap = new ColumnFieldMap(clazz);
		Workbook workbook = WorkbookUtils.save(templateInputStream, sheetName, columnFieldMap, list, rowStyle);
		workbook.write(outputStream);
		workbook.close();
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream object.
	 * @param templateInputStream InputStream for template.
	 * @param sheetName Sheet name in template and will in outputStream.
	 * @param list {@code List<T>} object.
	 * @param outputStream OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, List<T> list, OutputStream outputStream) throws Exception {
		save(templateInputStream, sheetName, list, null, outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path, save list data into it.
	 * @param templateInputStream InputStream for template.
	 * @param sheetName Sheet name in template and will in output Excel file.
	 * @param list {@code List<T>} object.
	 * @param rowStyle Style for row.
	 * @param filePath File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, List<T> list, RowStyle rowStyle, String filePath) throws Exception {
		FileOutputStream outputStream = new FileOutputStream(filePath, false);
		save(templateInputStream, sheetName, list, rowStyle, outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path, save list data into it.
	 * @param templateInputStream InputStream for template.
	 * @param sheetName Sheet name in template and will in output Excel file.
	 * @param list {@code List<T>} object.
	 * @param filePath File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(InputStream templateInputStream, String sheetName, List<T> list, String filePath) throws Exception {
		save(templateInputStream, sheetName, list, null, filePath);
	}

	/**
	 * Save the list data into specific OutputStream.
	 * @param excelType ExcelType object.
	 * @param sheetName Sheet name in outputStream.
	 * @param list {@code List<T>} object.
	 * @param outputStream OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, String sheetName, List<T> list, OutputStream outputStream) throws Exception {
		if (CollectionUtils.isNullOrEmpty(list)) {
			return;
		}

		Class<?> clazz = list.get(0).getClass();
		ColumnFieldMap columnFieldMap = new ColumnFieldMap(clazz);
		Workbook workbook = WorkbookUtils.save(excelType, sheetName, columnFieldMap, list);
		workbook.write(outputStream);
		workbook.close();
	}

	/**
	 * Save the list data into specific OutputStream.
	 * @param excelType ExcelType object.
	 * @param list {@code List<T>} object.
	 * @param outputStream OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(ExcelType excelType, List<T> list, OutputStream outputStream) throws Exception {
		String sheetName = SheetUtils.getName(1);
		save(excelType, sheetName, list, outputStream);
	}

	/**
	 * Create a new Excel file with specific excelType, sheet name and file path, save list data into it.
	 * @param sheetName Sheet name.
	 * @param list {@code List<T>} object.
	 * @param filePath File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(String sheetName, List<T> list, String filePath) throws Exception {
		FileOutputStream outputStream = new FileOutputStream(filePath, false);
		ExcelType excelType = ExcelUtils.getExcelType(filePath);
		save(excelType, sheetName, list, outputStream);
	}

	/**
	 * Create a new Excel file with specific excelType and file path, save list data into it.
	 * @param list {@code List<T>} object.
	 * @param filePath File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(List<T> list, String filePath) throws Exception {
		String sheetName = SheetUtils.getName(1);
		save(sheetName, list, filePath);
	}

	/**
	 * Append list data into template Excel file and save them into OutputStream object.
	 * @param templateWorkbook Workbook for template.
	 * @param sheetName Sheet name in template and will in outputStream.
	 * @param list {@code List<T>} object.
	 * @param rowStyle Style for row.
	 * @param outputStream OutputStream object that receive list data.
	 * @throws Exception
	 */
	public static <T> void save(Workbook templateWorkbook, String sheetName, List<T> list, RowStyle rowStyle, OutputStream outputStream) throws Exception {
		if (CollectionUtils.isNullOrEmpty(list)) {
			return;
		}

		Class<?> clazz = list.get(0).getClass();
		ColumnFieldMap columnFieldMap = new ColumnFieldMap(clazz);
		Workbook workbook = WorkbookUtils.save(templateWorkbook, sheetName, columnFieldMap, list, rowStyle);
		workbook.write(outputStream);
	}

	/**
	 * Create a new Excel file with specific template, sheet name and file path, save list data into it.
	 * @param templateWorkbook Workbook for template.
	 * @param sheetName Sheet name in template and will in output Excel file.
	 * @param list {@code List<T>} object.
	 * @param rowStyle Style for row.
	 * @param filePath File path for new Excel file.
	 * @throws Exception
	 */
	public static <T> void save(Workbook templateWorkbook, String sheetName, List<T> list, RowStyle rowStyle, String filePath) throws Exception {
		FileOutputStream outputStream = new FileOutputStream(filePath, false);
		save(templateWorkbook, sheetName, list, rowStyle, outputStream);
	}
}