package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.models.ExcelModel;
import devutility.external.poi.models.FieldColumnMap;
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
		FieldColumnMap<ExcelModel> fieldColumnMap = ExcelModel.getFieldColumnMap();
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", fieldColumnMap, list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void processExcel2007(int listCount) {
		String filePath = "E:\\Downloads\\Test.xlsx";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test.xlsx");
		FieldColumnMap<ExcelModel> fieldColumnMap = ExcelModel.getFieldColumnMap();
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			PoiUtils.save(templateInputStream, "Sheet1", fieldColumnMap, list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveWithTemplateTest.class);
	}
}