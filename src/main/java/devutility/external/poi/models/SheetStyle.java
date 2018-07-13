package devutility.external.poi.models;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class SheetStyle {
	/**
	 * Row Style.
	 */
	private CellStyle rowStyle;

	/**
	 * Row height.
	 */
	private float rowHeight;

	/**
	 * Row number.
	 */
	private int rowNum;

	/**
	 * Column style map.
	 */
	private Map<Integer, CellStyle> columnStyleMap;

	public SheetStyle(Workbook workbook, Row row) {
		rowStyle = row.getRowStyle();
		rowHeight = row.getHeightInPoints();

		short firstCellNum = row.getFirstCellNum();
		short lastCellNum = row.getLastCellNum();
		columnStyleMap = new HashMap<>(lastCellNum);

		for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
			Cell cell = row.getCell(cellNum);

			if (cell != null && cell.getCellStyle() != null) {
				columnStyleMap.put(cellNum, cell.getCellStyle());
			}
		}
	}

	public CellStyle setColumnStyle(int columnNum, CellStyle cellStyle) {
		return columnStyleMap.put(columnNum, cellStyle);
	}

	public CellStyle getColumnStyle(int columnNum) {
		return columnStyleMap.get(columnNum);
	}

	public CellStyle getRowStyle() {
		return rowStyle;
	}

	public void setRowStyle(CellStyle rowStyle) {
		this.rowStyle = rowStyle;
	}

	public float getRowHeight() {
		return rowHeight;
	}

	public void setRowHeight(float rowHeight) {
		this.rowHeight = rowHeight;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public Map<Integer, CellStyle> getColumnStyleMap() {
		return columnStyleMap;
	}

	public void setColumnStyleMap(Map<Integer, CellStyle> columnStyleMap) {
		this.columnStyleMap = columnStyleMap;
	}
}