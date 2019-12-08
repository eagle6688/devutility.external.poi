package devutility.external.poi.poiutils;

import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.model.ExcelModel;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveForBigDataTest extends BaseTest {
	@Override
	public void run() {
		List<ExcelModel> list = ExcelModel.create(100000);
		String filePath = "E:\\Downloads\\SaveForBigDataTest.xlsx";

		try {
			PoiUtils.save(list, filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveForBigDataTest.class);
	}
}