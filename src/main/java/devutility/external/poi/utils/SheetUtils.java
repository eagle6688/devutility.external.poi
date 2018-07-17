package devutility.external.poi.utils;

import java.lang.reflect.InvocationTargetException;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import devutility.external.poi.common.ExcelConfig;
import devutility.external.poi.models.FieldColumnEntry;
import devutility.external.poi.models.FieldColumnMap;
import devutility.external.poi.models.RowStyle;

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
			throw new Exception(String.format("Sheet %s not found!", name));
		}

		return sheet;
	}

	/**
	 * Get sheet name by sheet name format.
	 * @param index: Sheet index
	 * @return String
	 */
	public static String getName(int index) {
		return String.format(ExcelConfig.SHEETNAMEFORMAT, index);
	}

	/**
	 * Set new name for specific sheet.
	 * @param workbook: Workbook object.
	 * @param name: Sheet name.
	 * @param newName: Sheet new name.
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
	 * @param sheet: Sheet object.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List object.
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> void append(Sheet sheet, FieldColumnMap<T> fieldColumnMap, List<T> list) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		int rowNum = getStartRowNum(sheet);
		List<FieldColumnEntry> fieldColumnEntries = fieldColumnMap.getSortedEntries();

		if (list instanceof LinkedList) {
			for (T model : list) {
				RowUtils.create(sheet, rowNum++, model, fieldColumnEntries);
			}

			return;
		}

		for (int i = 0; i < list.size(); i++) {
			RowUtils.create(sheet, rowNum++, list.get(i), fieldColumnEntries);
		}
	}

	/**
	 * Append list to existed Sheet object.
	 * @param sheet: Sheet object.
	 * @param fieldColumnMap: FieldColumnMap object.
	 * @param list: List object.
	 * @param rowStyle: Style for each row.
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	public static <T> void append(Sheet sheet, FieldColumnMap<T> fieldColumnMap, List<T> list, RowStyle rowStyle) throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		int rowNum = getStartRowNum(sheet);
		List<FieldColumnEntry> fieldColumnEntries = fieldColumnMap.getSortedEntries();

		if (list instanceof LinkedList) {
			for (T model : list) {
				RowUtils.create(sheet, rowNum++, model, fieldColumnEntries, rowStyle);
			}

			return;
		}

		for (int i = 0; i < list.size(); i++) {
			RowUtils.create(sheet, rowNum++, list.get(i), fieldColumnEntries, rowStyle);
		}
	}

	public static int getStartRowNum(Sheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		Row row = sheet.getRow(lastRowNum);
		short firstCellNum = row.getFirstCellNum();

		if (firstCellNum > -1) {
			return lastRowNum + 1;
		}

		return lastRowNum;
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