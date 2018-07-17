package devutility.external.poi.utils;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import devutility.external.poi.models.RowStyle;

public class RowStyleUtils {
	public static RowStyle clone(Workbook workbook, String sheetName, int rowNum, boolean excludeBold) throws Exception {
		Sheet sheet = SheetUtils.get(workbook, sheetName);
		Row row = sheet.getRow(rowNum);
		return clone(workbook, row, excludeBold);
	}

	public static RowStyle clone(Workbook workbook, Row row, boolean excludeBold) {
		short firstCellNum = row.getFirstCellNum();
		short lastCellNum = row.getLastCellNum();

		if (firstCellNum == -1 || lastCellNum == -1) {
			return null;
		}

		Map<Integer, CellStyle> columnStyleMap = new HashMap<>(lastCellNum);

		/**
		 * Map for CellStyle number between cell style and new column style.
		 */
		Map<Short, Short> columnStyleIndexMap = new HashMap<>(lastCellNum);

		/**
		 * Map for Font number between cell style and new column style.
		 */
		Map<Short, Short> columnfontIndexMap = new HashMap<>(lastCellNum);

		for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
			Cell cell = row.getCell(cellNum);

			if (cell == null) {
				continue;
			}

			CellStyle cellStyle = cell.getCellStyle();

			if (cellStyle == null) {
				continue;
			}

			CellStyle newCellStyle = createCellStyle(workbook, columnStyleIndexMap, cellStyle);
			Font newFont = createFont(workbook, columnfontIndexMap, cellStyle);

			if (excludeBold) {
				newFont.setBold(false);
			}

			newCellStyle.setFont(newFont);
			columnStyleMap.put(cellNum, newCellStyle);
		}

		RowStyle rowStyle = new RowStyle();
		rowStyle.setRowNum(row.getRowNum());
		rowStyle.setRowHeight(row.getHeightInPoints());
		rowStyle.setRowStyle(row.getRowStyle());
		rowStyle.setColumnStyleMap(columnStyleMap);
		return rowStyle;
	}

	private static CellStyle createCellStyle(Workbook workbook, Map<Short, Short> columnStyleIndexMap, CellStyle cellStyle) {
		short cellStyleIndex = cellStyle.getIndex();
		Short newCellStyleIndex = columnStyleIndexMap.get(cellStyleIndex);

		if (newCellStyleIndex == null) {
			return cloneCellStyle(workbook, columnStyleIndexMap, cellStyle);
		}

		return workbook.getCellStyleAt(newCellStyleIndex);
	}

	private static CellStyle cloneCellStyle(Workbook workbook, Map<Short, Short> columnStyleIndexMap, CellStyle cellStyle) {
		CellStyle newCellStyle = CellStyleUtils.clone(workbook, cellStyle);
		short cellStyleIndex = cellStyle.getIndex();
		short newCellStyleIndex = newCellStyle.getIndex();
		columnStyleIndexMap.put(cellStyleIndex, newCellStyleIndex);
		return newCellStyle;
	}

	private static Font createFont(Workbook workbook, Map<Short, Short> columnfontIndexMap, CellStyle cellStyle) {
		short fontIndex = cellStyle.getFontIndex();
		Short newFontIndex = columnfontIndexMap.get(fontIndex);

		if (newFontIndex == null) {
			return cloneFont(workbook, columnfontIndexMap, fontIndex);
		}

		return workbook.getFontAt(newFontIndex);
	}

	private static Font cloneFont(Workbook workbook, Map<Short, Short> columnfontIndexMap, short fontIndex) {
		Font font = workbook.getFontAt(fontIndex);
		Font newFont = FontUtils.clone(workbook, font);
		short newFontIndex = newFont.getIndex();
		columnfontIndexMap.put(fontIndex, newFontIndex);
		return newFont;
	}
}