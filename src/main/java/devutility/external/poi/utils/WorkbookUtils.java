package devutility.external.poi.utils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import devutility.external.poi.common.ExcelType;
import devutility.external.poi.models.FieldColumnMap;

public class WorkbookUtils {
	/**
	 * Load excel file with file path.
	 * @param filePath: Excel file path.
	 * @return Workbook object.
	 * @throws EncryptedDocumentException
	 * @throws InvalidFormatException
	 * @throws IOException
	 */
	public static Workbook load(String filePath) throws IOException, EncryptedDocumentException, InvalidFormatException {
		File file = new File(filePath);

		if (!file.exists()) {
			throw new IOException(String.format("%s cannot found!", filePath));
		}

		return WorkbookFactory.create(file);
	}

	/**
	 * Create an Workbook instance use ExcelType.
	 * @param excelType: ExcelType object.
	 * @return Workbook
	 */
	public static Workbook create(ExcelType excelType) {
		if (excelType.equals(ExcelType.Excel2003)) {
			return new HSSFWorkbook();
		}

		return new XSSFWorkbook();
	}

	/**
	 * Save list data into specific sheet in template Workbook object.
	 * @param templateInputStream: Template InputStream.
	 * @param sheetName: Sheet name in template.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @return Workbook
	 * @throws Exception
	 */
	public static <T> Workbook save(InputStream templateInputStream, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list) throws Exception {
		Workbook templateWorkbook = WorkbookFactory.create(templateInputStream);
		return save(templateWorkbook, sheetName, fieldColumnMap, list);
	}

	/**
	 * Save list data into specific sheet in template Workbook object.
	 * @param templateWorkbook: Template Workbook object.
	 * @param sheetName: Sheet name in template.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @return Workbook
	 * @throws Exception
	 */
	public static <T> Workbook save(Workbook templateWorkbook, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list) throws Exception {
		Sheet sheet = SheetUtils.get(templateWorkbook, sheetName);
		SheetUtils.append(sheet, fieldColumnMap, list);
		return templateWorkbook;
	}

	/**
	 * Create a new Workbook instance with specific ExcelType and sheet name, save
	 * list data into it.
	 * @param excelType: ExcelType object.
	 * @param sheetName: Sheet name.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @return Workbook
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> Workbook save(ExcelType excelType, String sheetName, FieldColumnMap<T> fieldColumnMap, List<T> list) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Workbook workbook = WorkbookUtils.create(excelType);
		Sheet sheet = SheetUtils.create(workbook, sheetName);
		SheetUtils.append(sheet, fieldColumnMap, list);
		return workbook;
	}
}