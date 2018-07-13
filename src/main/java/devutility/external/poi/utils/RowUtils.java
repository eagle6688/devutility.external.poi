package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

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

	/**
	 * Create a new Row in specific Sheet.
	 * @param workbook: Workbook object.
	 * @param model: Model object.
	 * @param fieldColumnEntries: FieldColumnEntry list for type T.
	 * @return Row
	 * @throws InvocationTargetException
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
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
				row.createCell(entry.getColumnIndex()).setCellValue(value.toString());
			}
		}

		return row;
	}

	public static <T> Row clone(Row templateRow, Sheet sheet) {
		int rowNum = templateRow.getRowNum();
		short firstCellNum = templateRow.getFirstCellNum();
		short lastCellNum = templateRow.getLastCellNum();

		Row row = sheet.createRow(rowNum);
		row.setHeightInPoints(templateRow.getHeightInPoints());
		row.setRowStyle(templateRow.getRowStyle());
		row.setZeroHeight(templateRow.getZeroHeight());

		for (short cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
			
		}

		return row;
	}
}