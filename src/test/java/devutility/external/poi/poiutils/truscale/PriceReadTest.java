package devutility.external.poi.poiutils.truscale;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;
import devutility.internal.util.CollectionUtils;

public class PriceReadTest extends BaseTest {
	@Override
	public void run() {
		List<PriceVo> list = new LinkedList<PriceVo>();

		try (InputStream inputStream = new FileInputStream("E:\\Test\\PriceImportTemplate.xlsx")) {
			list = PoiUtils.read(inputStream, "Sheet1", PriceVo.class);
		} catch (Exception e) {
			e.printStackTrace();
		}

		if (CollectionUtils.isNotEmpty(list)) {
			list.forEach(i -> {
				System.out.println(i.toString());
			});
		}
	}

	public static void main(String[] args) {
		TestExecutor.run(PriceReadTest.class);
	}
}