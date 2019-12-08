package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.model.ExcelModel;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveWithTemplateTest extends BaseTest {
	@Override
	public void run() {
		int count = 100;
		processExcel2003(count);
		processExcel2007(count);
	}

	private void processExcel2003(int listCount) {
		String filePath = "E:\\Downloads\\Test.xls";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test.xls");
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void processExcel2007(int listCount) {
		String filePath = "E:\\Downloads\\Test.xlsx";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test.xlsx");
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveWithTemplateTest.class);
	}
}