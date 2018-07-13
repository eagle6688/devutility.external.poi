package devutility.external.poi.utils;

import java.io.File;
import java.io.IOException;
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
	 * Clone then templated Sheet in templated Workbook, rename it with specific
	 * sheet name, save list data into it.
	 * @param templateWorkbook: Template Workbook object.
	 * @param templateSheetName: Sheet name in template.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List data.
	 * @param sheetName: New sheet name.
	 * @return Workbook
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	public static <T> Workbook save(Workbook templateWorkbook, String templateSheetName, FieldColumnMap<T> fieldColumnMap, List<T> list, String sheetName)
			throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		int templateSheetIndex = templateWorkbook.getSheetIndex(templateSheetName);

		if (templateSheetIndex == -1) {
			throw new NullPointerException(String.format("Sheet %s not found!", templateSheetName));
		}

		Sheet sheet = templateWorkbook.cloneSheet(templateSheetIndex);
		SheetUtils.setName(templateWorkbook, sheet.getSheetName(), sheetName);
		SheetUtils.append(sheet, fieldColumnMap, list);
		return templateWorkbook;
	}

	

}