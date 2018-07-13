package devutility.external.poi.workbook;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.utils.FontUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class CellStyleAndFontTest extends BaseTest {
	@Override
	public void run() {
		process("Test.xlsx");
	}

	void process(String file) {
		InputStream inputStream = CellStyleAndFontTest.class.getClassLoader().getResourceAsStream(file);
		Workbook workbook = null;

		try {
			workbook = WorkbookFactory.create(inputStream);
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
		}

		println(workbook.getNumberOfFonts());
		println(workbook.getNumCellStyles());

		for (int i = 0; i < workbook.getNumberOfFonts(); i++) {
			Font font = workbook.getFontAt((short) i);
			println(font.getFontName());
			println(String.valueOf(font.getBold()));
		}

		if (workbook.getFontAt((short) 0).equals(workbook.getFontAt((short) 1))) {
			println("Equal");
		} else {
			println("Not equal");
		}

		println("===============New===============");
		Font font0 = workbook.getFontAt((short) 0);
		Font newFont = FontUtils.clone(workbook, font0);
		println(workbook.getNumberOfFonts());

		if (newFont.equals(font0)) {
			println("Equal");
		} else {
			println("Not equal");
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(CellStyleAndFontTest.class);
	}
}