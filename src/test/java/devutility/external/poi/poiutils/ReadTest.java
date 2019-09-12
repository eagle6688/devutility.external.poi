package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.models.ExcelModel;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class ReadTest extends BaseTest {
	@Override
	public void run() {
		process("Test.xls");
		process("Test.xlsx");
	}

	private void process(String fileName) {
		try (InputStream inputStream = ReadTest.class.getClassLoader().getResourceAsStream(fileName)) {
			if (inputStream == null) {
				println(String.format("%s not found!", fileName));
				return;
			}

			List<ExcelModel> list = PoiUtils.read(inputStream, "Sheet1", ExcelModel.class);

			for (ExcelModel test : list) {
				println(test.toString());
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(ReadTest.class);
	}
}