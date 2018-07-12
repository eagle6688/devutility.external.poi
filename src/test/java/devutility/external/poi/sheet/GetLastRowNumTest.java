package devutility.external.poi.sheet;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.common.ExcelType;
import devutility.external.poi.utils.WorkbookUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class GetLastRowNumTest extends BaseTest {
	@Override
	public void run() {
		process(ExcelType.Excel2003);
		process(ExcelType.Excel2007);

		try {
			process("Test.xls", "Sheet1");
			process("Test.xlsx", "Sheet1");
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}
	}

	private void process(ExcelType excelType) {
		Workbook workbook = WorkbookUtils.create(excelType);
		Sheet sheet = workbook.createSheet("Sheet1");
		println(sheet.getLastRowNum());
	}

	private void process(String file, String sheet) throws EncryptedDocumentException, InvalidFormatException, IOException {
		InputStream inputStream = GetLastRowNumTest.class.getClassLoader().getResourceAsStream(file);
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheetObj = workbook.getSheet(sheet);
		println(sheetObj.getLastRowNum());
	}

	public static void main(String[] args) {
		TestExecutor.run(GetLastRowNumTest.class);
	}
}