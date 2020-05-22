package devutility.external.poi.model;

import devutility.external.poi.common.ExcelColumn;
import devutility.internal.test.BaseTest;
import devutility.internal.test.TestExecutor;

public class ColumnFieldMapTest extends BaseTest {
	@Override
	public void run() {
		new ColumnFieldMap(ColumnFieldMapTestModel.class);
	}

	public static void main(String[] args) {
		TestExecutor.run(ColumnFieldMapTest.class);
	}
}

class ColumnFieldMapTestModel {
	@ExcelColumn(index = 0)
	private String field1;

	private String field2;

	@ExcelColumn(index = 1)
	private String field3;

	public String getField1() {
		return field1;
	}

	public void setField1(String field1) {
		this.field1 = field1;
	}

	public String getField2() {
		return field2;
	}

	public void setField2(String field2) {
		this.field2 = field2;
	}

	public String getField3() {
		return field3;
	}

	public void setField3(String field3) {
		this.field3 = field3;
	}
}