package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import devutility.external.poi.common.ExcelConfig;
import devutility.external.poi.models.ColumnFieldMap;
import devutility.external.poi.models.RowStyle;

public class SheetUtils {
	/**
	 * Create an Sheet object with unsafe name.
	 * @param workbook Workbook object.
	 * @param name Sheet name.
	 * @return Sheet
	 */
	public static Sheet create(Workbook workbook, String name) {
		String safeName = WorkbookUtil.createSafeSheetName(name);
		return workbook.createSheet(safeName);
	}

	/**
	 * Get an Sheet object.
	 * @param workbook Workbook object.
	 * @param name Sheet name.
	 * @return Sheet
	 * @throws Exception
	 */
	public static Sheet get(Workbook workbook, String name) throws Exception {
		Sheet sheet = workbook.getSheet(name);

		if (sheet == null) {
			throw new Exception(String.format("Sheet %s not found!", name));
		}

		return sheet;
	}

	/**
	 * Get sheet name by sheet name format.
	 * @param index Sheet index.
	 * @return String
	 */
	public static String getName(int index) {
		return String.format(ExcelConfig.SHEETNAMEFORMAT, index);
	}

	/**
	 * Set new name for specific sheet.
	 * @param workbook Workbook object.
	 * @param name Sheet name.
	 * @param newName Sheet new name.
	 */
	public static void setName(Workbook workbook, String name, String newName) {
		int sheetIndex = workbook.getSheetIndex(name);

		if (sheetIndex == -1) {
			throw new NullPointerException(String.format("Sheet %s not found!", name));
		}

		String safeName = WorkbookUtil.createSafeSheetName(newName);
		workbook.setSheetName(sheetIndex, safeName);
	}

	/**
	 * Append list to existed Sheet object.
	 * @param sheet Sheet object.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param list List object.
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> void append(Sheet sheet, ColumnFieldMap columnFieldMap, List<T> list) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		int rowNum = getStartRowNum(sheet);

		for (int i = 0; i < list.size(); i++) {
			RowUtils.create(sheet, rowNum++, columnFieldMap, list.get(i));
		}
	}

	/**
	 * Append list to existed Sheet object.
	 * @param sheet Sheet object.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param list {@code List<T>} object.
	 * @param rowStyle Style for each row.
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> void append(Sheet sheet, ColumnFieldMap columnFieldMap, List<T> list, RowStyle rowStyle) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		int rowNum = getStartRowNum(sheet);

		for (int i = 0; i < list.size(); i++) {
			RowUtils.create(sheet, rowNum++, columnFieldMap, list.get(i), rowStyle);
		}
	}

	/**
	 * Return start row number.
	 * @param sheet Sheet object.
	 * @return int
	 */
	public static int getStartRowNum(Sheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);

		if (lastRowNum == 0 && row == null) {
			return 0;
		}

		short firstCellNum = row.getFirstCellNum();

		if (firstCellNum > -1) {
			return lastRowNum + 1;
		}

		return lastRowNum;
	}

	/**
	 * Convert an Sheet object to list(type T).
	 * @param sheet Sheet object.
	 * @param columnFieldMap The map between Excel column index and type T field.
	 * @param clazz Class object.
	 * @return {@code List<T>}
	 * @throws ReflectiveOperationException
	 */
	public static <T> List<T> toList(Sheet sheet, ColumnFieldMap columnFieldMap, Class<T> clazz) throws ReflectiveOperationException {
		List<T> list = new LinkedList<>();

		if (sheet == null) {
			return list;
		}

		int rowStart = sheet.getFirstRowNum();
		int rowEnd = sheet.getLastRowNum() + 1;

		for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
			Row row = sheet.getRow(rowNum);

			if (row == null) {
				continue;
			}

			T model = RowUtils.toModel(row, columnFieldMap, clazz);

			if (model != null) {
				list.add(model);
			}
		}

		return list;
	}
}