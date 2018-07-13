package devutility.external.poi.poiutils;

import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.models.ExcelModel;
import devutility.external.poi.models.FieldColumnMap;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveForBigDataTest extends BaseTest {
	@Override
	public void run() {
		String filePath = "E:\\Downloads\\Test.xlsx";
		FieldColumnMap<ExcelModel> fieldColumnMap = ExcelModel.getFieldColumnMap();
		List<ExcelModel> list = ExcelModel.create(100000);

		try {
			PoiUtils.save(fieldColumnMap, list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveForBigDataTest.class);
	}
}