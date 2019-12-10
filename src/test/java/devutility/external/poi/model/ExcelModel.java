package devutility.external.poi.model;

import java.util.ArrayList;
import java.util.List;

import devutility.external.poi.common.ExcelColumn;

public class ExcelModel {
	@ExcelColumn(index = 0)
	private String first;

	@ExcelColumn(index = 2)
	private String second;

	@ExcelColumn(index = 1)
	private String third;

	private String fifth;

	@ExcelColumn(index = 4)
	private String fourth;

	public ExcelModel() {

	}

	public ExcelModel(String first, String second, String third, String fourth, String fifth) {
		this.first = first;
		this.second = second;
		this.third = third;
		this.fourth = fourth;
		this.fifth = fifth;
	}

	public String toString() {
		return String.format("First: %s, Second: %s, Third: %s, Fourth: %s, Fifth: %s", first, second, third, fourth, fifth);
	}

	public static List<ExcelModel> create(int count) {
		List<ExcelModel> list = new ArrayList<>(count);

		for (int i = 0; i < count; i++) {
			int num = i + 1;
			String first = String.format("First-%d", num);
			String second = String.format("Second-%d", num);
			String third = String.format("Third-%d", num);
			String fourth = String.format("Fourth-%d", num);
			String fifth = String.format("Fifth-%d", num);
			ExcelModel excelModel = new ExcelModel(first, second, third, fourth, fifth);
			list.add(excelModel);
		}

		return list;
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

	public String getFifth() {
		return fifth;
	}

	public void setFifth(String fifth) {
		this.fifth = fifth;
	}

	public String getFourth() {
		return fourth;
	}

	public void setFourth(String fourth) {
		this.fourth = fourth;
	}
}