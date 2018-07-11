package devutility.external.poi.models;

import devutility.internal.lang.models.EntityField;

public class FieldColumnEntry {
	private int columnIndex;
	private String modelField;
	private EntityField entityField;

	public FieldColumnEntry() {
	}

	public FieldColumnEntry(int columnIndex, String modelField) {
		this.columnIndex = columnIndex;
		this.modelField = modelField;
	}

	public int getColumnIndex() {
		return columnIndex;
	}

	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}

	public String getModelField() {
		return modelField;
	}

	public void setModelField(String modelField) {
		this.modelField = modelField;
	}

	public EntityField getEntityField() {
		return entityField;
	}

	public void setEntityField(EntityField entityField) {
		this.entityField = entityField;
	}
}