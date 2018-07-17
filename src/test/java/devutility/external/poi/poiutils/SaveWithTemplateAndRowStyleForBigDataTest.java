package devutility.external.poi.poiutils;

import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import devutility.external.poi.PoiUtils;
import devutility.external.poi.models.ExcelModel;
import devutility.external.poi.models.FieldColumnMap;
import devutility.external.poi.models.RowStyle;
import devutility.external.poi.utils.RowStyleUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class SaveWithTemplateAndRowStyleForBigDataTest extends BaseTest {
	@Override
	public void run() {
		processExcel2007(1000);
	}

	private void processExcel2007(int listCount) {
		String filePath = "E:\\Downloads\\Test.xlsx";
		InputStream templateInputStream = SaveWithTemplateTest.class.getClassLoader().getResourceAsStream("Test.xlsx");
		FieldColumnMap<ExcelModel> fieldColumnMap = ExcelModel.getFieldColumnMap();
		List<ExcelModel> list = ExcelModel.create(listCount);

		try {
			Workbook workbook = WorkbookFactory.create(templateInputStream);
			RowStyle rowStyle = RowStyleUtils.clone(workbook, "Sheet1", 0, true);
			PoiUtils.save(workbook, "Sheet1", fieldColumnMap, list, rowStyle, filePath);
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(SaveWithTemplateAndRowStyleForBigDataTest.class);
	}
}