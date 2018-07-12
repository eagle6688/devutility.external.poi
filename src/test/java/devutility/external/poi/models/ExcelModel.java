package devutility.external.poi.models;

import java.util.ArrayList;
import java.util.List;

public class ExcelModel {
	private String first;
	private String second;
	private String third;
	private String fourth;

	public ExcelModel() {

	}

	public ExcelModel(String first, String second, String third, String fourth) {
		this.first = first;
		this.second = second;
		this.third = third;
		this.fourth = fourth;
	}

	public String toString() {
		return String.format("First: %s, Second: %s, Third: %s, Fourth: %s", first, second, third, fourth);
	}

	public static List<ExcelModel> create(int count) {
		List<ExcelModel> list = new ArrayList<>(count);

		for (int i = 0; i < count; i++) {
			int num = i + 1;
			String first = String.format("First-%d", num);
			String second = String.format("Second-%d", num);
			String third = String.format("Third-%d", num);
			String fourth = String.format("Fourth-%d", num);
			ExcelModel excelModel = new ExcelModel(first, second, third, fourth);
			list.add(excelModel);
		}

		return list;
	}

	public static FieldColumnMap<ExcelModel> getFieldColumnMap() {
		FieldColumnMap<ExcelModel> fieldColumnMap = new FieldColumnMap<>(ExcelModel.class);
		fieldColumnMap.put(0, "first");
		fieldColumnMap.put(1, "second");
		fieldColumnMap.put(2, "third");
		fieldColumnMap.put(3, "fourth");
		return fieldColumnMap;
	}

	public String getFirst() {
		return first;
	}

	public void setFirst(String first) {
		this.first = first;
	}

	public String getSecond() {
		return second;
	}

	public void setSecond(String second) {
		this.second = second;
	}

	public String getThird() {
		return third;
	}

	public void setThird(String third) {
		this.third = third;
	}

	public String getFourth() {
		return fourth;
	}

	public void setFourth(String fourth) {
		this.fourth = fourth;
	}
}