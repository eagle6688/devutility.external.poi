package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.models.ExcelModel;
import devutility.external.poi.models.FieldColumnMap;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveWithTemplate1ForBigDataTest extends BaseTest {
	@Override
	public void run() {
		processExcel2007(10000);
	}

	private void processExcel2007(int listCount) {
		String filePath = "E:\\Downloads\\SaveWithTemplate1ForBigDataTest.xlsx";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test1.xlsx");
		FieldColumnMap<ExcelModel> fieldColumnMap = ExcelModel.getFieldColumnMap();
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", fieldColumnMap, list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveWithTemplate1ForBigDataTest.class);
	}
}