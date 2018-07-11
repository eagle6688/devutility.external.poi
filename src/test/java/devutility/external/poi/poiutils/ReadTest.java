package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.Test;
import devutility.external.poi.models.FieldColumnMap;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class ReadTest extends BaseTest {
	@Override
	public void run() {
		process("Test.xls");
		process("Test.xlsx");
	}

	private void process(String fileName) {
		FieldColumnMap<Test> fieldColumnMap = new FieldColumnMap<>(Test.class);
		fieldColumnMap.put(0, "first");
		fieldColumnMap.put(1, "second");
		fieldColumnMap.put(2, "third");
		fieldColumnMap.put(3, "fourth");

		try (InputStream inputStream = ReadTest.class.getClassLoader().getResourceAsStream(fileName)) {
			if (inputStream == null) {
				println(String.format("%s not found!", fileName));
				return;
			}

			List<Test> list = PoiUtils.read(inputStream, "Sheet1", fieldColumnMap, Test.class);

			for (Test test : list) {
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