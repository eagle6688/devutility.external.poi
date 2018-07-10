package devutility.external.poi.utils;

import java.lang.reflect.Method;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import devutility.external.poi.models.FieldColumnEntry;
import devutility.external.poi.models.FieldColumnMap;
import devutility.internal.lang.ClassHelper;
import devutility.internal.lang.models.EntityField;

public class RowUtils {
	/**
	 * Convert row to T type model.
	 * @param row: Row object.
	 * @param dataFormatter: DataFormatter object.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param clazz: Class object.
	 * @return {@code T}
	 * @throws ReflectiveOperationException
	 */
	public static <T> T toModel(Row row, DataFormatter dataFormatter, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws ReflectiveOperationException {
		T model = ClassHelper.newInstance(clazz);

		for (FieldColumnEntry entry : fieldColumnMap.getEntries()) {
			EntityField entityField = entry.getEntityField();

			if (entityField == null) {
				continue;
			}

			Method setter = entityField.getSetter();

			if (setter == null) {
				continue;
			}

			int columnIndex = entry.getColumnIndex();
			Cell cell = row.getCell(columnIndex);
			Object value = CellUtils.getValue(cell, dataFormatter);

			if (value != null) {
				setter.invoke(model, value);
			}
		}

		return model;
	}
}