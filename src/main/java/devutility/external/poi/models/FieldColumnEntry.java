package devutility.external.poi.models;

public class FieldColumnEntry {
	private String modelField;
	private int columnIndex;

	public FieldColumnEntry() {
	}

	public FieldColumnEntry(String modelField, int columnIndex) {
		this.modelField = modelField;
		this.columnIndex = columnIndex;
	}

	public String getModelField() {
		return modelField;
	}

	public void setModelField(String modelField) {
		this.modelField = modelField;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}
}