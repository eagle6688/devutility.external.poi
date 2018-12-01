package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import devutility.external.poi.models.FieldColumnEntry;
import devutility.external.poi.models.FieldColumnMap;
import devutility.external.poi.models.RowStyle;
import devutility.internal.lang.ClassUtils;
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
		T model = ClassUtils.newInstance(clazz);

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

	/**
	 * Create a new Row in specific Sheet.
	 * @param sheet: Sheet object.
	 * @param rowNum: Row number.
	 * @param model: Model object.
	 * @param fieldColumnEntries: FieldColumnEntry list for type T.
	 * @return Row
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> Row create(Sheet sheet, int rowNum, T model, List<FieldColumnEntry> fieldColumnEntries) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Row row = sheet.createRow(rowNum);

		for (FieldColumnEntry entry : fieldColumnEntries) {
			Method getter = entry.getEntityField().getGetter();

			if (getter == null) {
				continue;
			}

			Object value = getter.invoke(model);

			if (value != null) {
				Cell cell = row.createCell(entry.getColumnIndex());
				cell.setCellValue(value.toString());
			}
		}

		return row;
	}

	/**
	 * Create a new Row in specific Sheet.
	 * @param sheet: Sheet object.
	 * @param rowNum: Row number.
	 * @param model: Model object.
	 * @param fieldColumnEntries: FieldColumnEntry list for type T.
	 * @param rowStyle: Style for whole row.
	 * @return Row
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> Row create(Sheet sheet, int rowNum, T model, List<FieldColumnEntry> fieldColumnEntries, RowStyle rowStyle) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		Row row = sheet.createRow(rowNum);

		for (FieldColumnEntry entry : fieldColumnEntries) {
			Method getter = entry.getEntityField().getGetter();

			if (getter == null) {
				continue;
			}

			Object value = getter.invoke(model);

			if (value != null) {
				Cell cell = row.createCell(entry.getColumnIndex());
				cell.setCellValue(value.toString());

				CellStyle cellStyle = rowStyle.getColumnStyle(entry.getColumnIndex());

				if (cellStyle != null) {
					cell.setCellStyle(cellStyle);
				}
			}
		}

		applyStyle(row, rowStyle);
		return row;
	}

	public static void applyStyle(Row row, RowStyle rowStyle) {
		if (rowStyle.getRowHeight() > 0) {
			row.setHeightInPoints(rowStyle.getRowHeight());
		}

		if (rowStyle.getRowStyle() != null) {
			row.setRowStyle(rowStyle.getRowStyle());
		}
	}
}