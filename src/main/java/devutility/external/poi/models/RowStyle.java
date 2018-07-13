package devutility.external.poi.models;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import devutility.external.poi.utils.CellStyleUtils;
import devutility.external.poi.utils.FontUtils;

public class RowStyle {
	/**
	 * Row number.
	 */
	private int rowNum;

	/**
	 * Row height.
	 */
	private float rowHeight;

	/**
	 * Row Style.
	 */
	private CellStyle rowStyle;

	/**
	 * Map for CellStyle number between cell style and new column style.
	 */
	private Map<Short, Short> columnStyleIndexMap;

	/**
	 * Map for Font number between cell style and new column style.
	 */
	private Map<Short, Short> fontIndexMap;

	/**
	 * Column style map.
	 */
	private Map<Integer, CellStyle> columnStyleMap;

	/**
	 * Constructor
	 * @param workbook: Workbook object.
	 * @param row: Row object.
	 */
	public RowStyle(Workbook workbook, Row row) {
		init(workbook, row);
	}

	private void init(Workbook workbook, Row row) {
		this.rowNum = row.getRowNum();
		this.rowHeight = row.getHeightInPoints();
		this.rowStyle = row.getRowStyle();

		short firstCellNum = row.getFirstCellNum();
		short lastCellNum = row.getLastCellNum();

		columnStyleIndexMap = new HashMap<>(lastCellNum);
		fontIndexMap = new HashMap<>(lastCellNum);
		columnStyleMap = new HashMap<>(lastCellNum);

		for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
			Cell cell = row.getCell(cellNum);

			if (cell == null) {
				continue;
			}

			CellStyle cellStyle = cell.getCellStyle();

			if (cellStyle == null) {
				continue;
			}

			CellStyle newCellStyle = saveCellStyle(workbook, cellStyle);
			Font newFont = saveFont(workbook, cellStyle);
			newCellStyle.setFont(newFont);

			columnStyleMap.put(cellNum, newCellStyle);
		}
	}

	private CellStyle saveCellStyle(Workbook workbook, CellStyle cellStyle) {
		short cellStyleIndex = cellStyle.getIndex();
		Short newCellStyleIndex = columnStyleIndexMap.get(cellStyleIndex);

		if (newCellStyleIndex == null) {
			return cloneCellStyle(workbook, cellStyle);
		}

		return workbook.getCellStyleAt(newCellStyleIndex);
	}

	private CellStyle cloneCellStyle(Workbook workbook, CellStyle cellStyle) {
		CellStyle newCellStyle = CellStyleUtils.clone(workbook, cellStyle);
		short cellStyleIndex = cellStyle.getIndex();
		short newCellStyleIndex = newCellStyle.getIndex();
		columnStyleIndexMap.put(cellStyleIndex, newCellStyleIndex);
		return newCellStyle;
	}

	private Font saveFont(Workbook workbook, CellStyle cellStyle) {
		short fontIndex = cellStyle.getFontIndex();
		Short newFontIndex = fontIndexMap.get(fontIndex);

		if (newFontIndex == null) {
			return cloneFont(workbook, fontIndex);
		}

		return workbook.getFontAt(newFontIndex);
	}

	private Font cloneFont(Workbook workbook, short fontIndex) {
		Font font = workbook.getFontAt(fontIndex);
		Font newFont = FontUtils.clone(workbook, font);
		short newFontIndex = newFont.getIndex();
		fontIndexMap.put(fontIndex, newFontIndex);
		return newFont;
	}

	public CellStyle setColumnStyle(int columnNum, CellStyle cellStyle) {
		return columnStyleMap.put(columnNum, cellStyle);
	}

	public CellStyle getColumnStyle(int columnNum) {
		return columnStyleMap.get(columnNum);
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public float getRowHeight() {
		return rowHeight;
	}

	public void setRowHeight(float rowHeight) {
		this.rowHeight = rowHeight;
	}

	public CellStyle getRowStyle() {
		return rowStyle;
	}

	public void setRowStyle(CellStyle rowStyle) {
		this.rowStyle = rowStyle;
	}

	public Map<Integer, CellStyle> getColumnStyleMap() {
		return columnStyleMap;
	}

	public void setColumnStyleMap(Map<Integer, CellStyle> columnStyleMap) {
		this.columnStyleMap = columnStyleMap;
	}
}