package devutility.external.poi.workbook;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.utils.FontUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class FontTest extends BaseTest {
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

		println(workbook.getNumberOfFontsAsInt());

		for (int i = 0; i < workbook.getNumberOfFontsAsInt(); i++) {
			Font font = workbook.getFontAt(i);
			println(font.getFontName());
			println(String.valueOf(font.getBold()));
		}

		if (workbook.getFontAt(0).equals(workbook.getFontAt(1))) {
			println("Equal");
		} else {
			println("Not equal");
		}

		println("===============New===============");
		Font font0 = workbook.getFontAt(0);
		Font newFont = FontUtils.clone(workbook, font0);
		println(workbook.getNumberOfFontsAsInt());

		if (newFont.equals(font0)) {
			println("Equal");
		} else {
			println("Not equal");
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(FontTest.class);
	}
}