package devutility.external.poi.workbook;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class CellStyleTest extends BaseTest {
	@Override
	public void run() {
		process("Test.xlsx");
	}

	void process(String file) {
		InputStream inputStream = FontTest.class.getClassLoader().getResourceAsStream(file);
		Workbook workbook = null;

		try {
			workbook = WorkbookFactory.create(inputStream);
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
		}

		Sheet sheet = workbook.getSheetAt(0);
		Row row = sheet.getRow(0);
		int count = row.getLastCellNum();

		for (int i = 0; i < count; i++) {
			Cell cell = row.getCell(i);
			println(cell.getCellStyle().getIndex());
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(CellStyleTest.class);
	}
}