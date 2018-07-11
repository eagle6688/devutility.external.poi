package devutility.external.poi;

public class Test {
	private String first;
	private String second;
	private String third;
	private String fourth;

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

	public String toString() {
		return String.format("First: %s, Second: %s, Third: %s, Fourth: %s", first, second, third, fourth);
	}
}