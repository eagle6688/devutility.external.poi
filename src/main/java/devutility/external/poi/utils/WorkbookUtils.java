package devutility.external.poi.utils;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import devutility.external.poi.common.ExcelType;

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
}