package devutility.external.poi.utils;

import java.util.LinkedList;
import java.util.List;

import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;

import devutility.external.poi.models.FieldColumnMap;

public class SheetUtils {

	/**
	 * Create a Sheet object with unsafe name.
	 * @param workbook: Workbook object.
	 * @param name: Sheet name.
	 * @return Sheet
	 */
	public static Sheet createSheet(Workbook workbook, String name) {
		String safeName = WorkbookUtil.createSafeSheetName(name);
		return workbook.createSheet(safeName);
	}

	public static Sheet getSheet(Workbook workbook, String name) throws Exception {
		Sheet sheet = workbook.getSheet(name);

		if (sheet == null) {
			throw new Exception(String.format("Sheet %s cannot found!", name));
		}

		return sheet;
	}

	/**
	 * Convert Sheet to list(type T)
	 * @param sheet
	 * @param fieldColumnMap
	 * @param clazz
	 * @return {@code List<T>}
	 * @throws ReflectiveOperationException
	 */
	public static List<T> toList(Sheet sheet, FieldColumnMap<T> fieldColumnMap, Class<T> clazz) throws ReflectiveOperationException {
		List<T> list = new LinkedList<>();

		if (sheet == null) {
			return list;
		}

		DataFormatter dataFormatter = new DataFormatter();
		int rowStart = sheet.getFirstRowNum();
		int rowEnd = sheet.getLastRowNum();

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