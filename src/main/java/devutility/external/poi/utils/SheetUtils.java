package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import devutility.external.poi.models.FieldColumnMap;

public class SheetUtils {
	/**
	 * Create an Sheet object with unsafe name.
	 * @param workbook: Workbook object.
	 * @param name: Sheet name.
	 * @return Sheet
	 */
	public static Sheet create(Workbook workbook, String name) {
		String safeName = WorkbookUtil.createSafeSheetName(name);
		return workbook.createSheet(safeName);
	}

	/**
	 * Get an Sheet object.
	 * @param workbook: Workbook object.
	 * @param name: Sheet name.
	 * @return Sheet
	 * @throws Exception
	 */
	public static Sheet get(Workbook workbook, String name) throws Exception {
		Sheet sheet = workbook.getSheet(name);

		if (sheet == null) {
			throw new Exception(String.format("Sheet %s cannot found!", name));
		}

		return sheet;
	}

	/**
	 * Append list to existed Sheet object.
	 * @param sheet: Sheet object.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List object.
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> void append(Sheet sheet, FieldColumnMap<T> fieldColumnMap, List<T> list) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		for (T model : list) {
			int rowNum = sheet.getLastRowNum() + 1;
			RowUtils.create(sheet, rowNum, model, fieldColumnMap);
		}
	}

	/**
	 * Convert an Sheet object to list(type T).
	 * @param sheet
	 * @param fieldColumnMap
	 * @param clazz
	 * @return {@code List<T>}
	 * @throws ReflectiveOperationException
	 */
	public static <T> List<T> toList(Sheet sheet, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws ReflectiveOperationException {
		List<T> list = new LinkedList<>();

		if (sheet == null) {
			return list;
		}

		DataFormatter dataFormatter = new DataFormatter();
		int rowStart = sheet.getFirstRowNum();
		int rowEnd = sheet.getLastRowNum() + 1;

		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = sheet.getRow(rowNum);

			if (row == null) {
				continue;
			}

			T model = RowUtils.toModel(row, dataFormatter, fieldColumnMap, clazz);

			if (model != null) {
				list.add(model);
			}
		}

		return list;
	}
}