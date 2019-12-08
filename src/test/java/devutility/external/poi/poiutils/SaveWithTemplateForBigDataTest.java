package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.model.ExcelModel;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveWithTemplateForBigDataTest extends BaseTest {
	@Override
	public void run() {
		processExcel2007(100000);
	}

	private void processExcel2007(int listCount) {
		String filePath = "E:\\Downloads\\SaveWithTemplateForBigDataTest.xlsx";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test.xlsx");
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveWithTemplateForBigDataTest.class);
	}
}