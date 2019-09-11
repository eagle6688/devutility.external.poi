package devutility.external.poi.poiutils.truscale;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

import devutility.external.poi.PoiUtils;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;
import devutility.internal.util.CollectionUtils;

public class ContractTest extends BaseTest {
	@Override
	public void run() {
		List<ContractVo> list = new LinkedList<ContractVo>();

		try (InputStream inputStream = new FileInputStream("E:\\Test\\contract.xlsx")) {
			list = PoiUtils.read(inputStream, "Sheet1", ContractVo.getFieldColumnMap(), ContractVo.class);
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
		TestExecutor.run(ContractTest.class);
	}
}