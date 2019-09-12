package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import devutility.external.poi.models.ColumnFieldMap;
import devutility.external.poi.models.RowStyle;
import devutility.internal.lang.ClassUtils;
import devutility.internal.lang.models.EntityField;

public class RowUtils {
	/**
	 * Convert row to T type model.
	 * @param row Row object.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param clazz Class object.
	 * @return {@code T}
	 * @throws ReflectiveOperationException
	 */
	public static <T> T toModel(Row row, ColumnFieldMap columnFieldMap, Class<T> clazz) throws ReflectiveOperationException {
		T model = ClassUtils.newInstance(clazz);

		for (Map.Entry<Integer, EntityField> entry : columnFieldMap.entrySet()) {
			EntityField entityField = entry.getValue();

			if (entityField == null) {
				continue;
			}

			Method setter = entityField.getSetter();

			if (setter == null) {
				continue;
			}

			int columnIndex = entry.getKey();
			Cell cell = row.getCell(columnIndex);
			Object value = CellUtils.getValue(cell, entityField.fieldType());

			if (value != null) {
				setter.invoke(model, value);
			}
		}

		return model;
	}

	/**
	 * Create a new Row in specific Sheet.
	 * @param sheet Sheet object.
	 * @param rowNum Row number.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param model Model object.
	 * @return Row
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> Row create(Sheet sheet, int rowNum, ColumnFieldMap columnFieldMap, T model) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Row row = sheet.createRow(rowNum);

		for (Map.Entry<Integer, EntityField> entry : columnFieldMap.entrySet()) {
			Method getter = entry.getValue().getGetter();

			if (getter == null) {
				continue;
			}

			Object value = getter.invoke(model);

			if (value != null) {
				Cell cell = row.createCell(entry.getKey());
				cell.setCellValue(value.toString());
			}
		}

		return row;
	}

	/**
	 * Create a new Row in specific Sheet.
	 * @param sheet Sheet object.
	 * @param rowNum Row number.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param model Model object.
	 * @param rowStyle Style for whole row.
	 * @return Row
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> Row create(Sheet sheet, int rowNum, ColumnFieldMap columnFieldMap, T model, RowStyle rowStyle) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Row row = sheet.createRow(rowNum);

		for (Map.Entry<Integer, EntityField> entry : columnFieldMap.entrySet()) {
			Method getter = entry.getValue().getGetter();

			if (getter == null) {
				continue;
			}

			Object value = getter.invoke(model);

			if (value != null) {
				Cell cell = row.createCell(entry.getKey());
				cell.setCellValue(value.toString());

				CellStyle cellStyle = rowStyle.getColumnStyle(entry.getKey());

				if (cellStyle != null) {
					cell.setCellStyle(cellStyle);
				}
			}
		}

		applyStyle(row, rowStyle);
		return row;
	}

	/**
	 * Apply provided style to provided row.
	 * @param row Row object
	 * @param rowStyle RowStyle object.
	 */
	public static void applyStyle(Row row, RowStyle rowStyle) {
		if (rowStyle.getRowHeight() > 0) {
			row.setHeightInPoints(rowStyle.getRowHeight());
		}

		if (rowStyle.getRowStyle() != null) {
			row.setRowStyle(rowStyle.getRowStyle());
		}
	}
}