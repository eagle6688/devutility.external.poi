package devutility.external.poi.model;

import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 
 * RowStyle
 * 
 * @author: Aldwin Su
 * @version: 2019-12-04 23:01:13
 */
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
	 * Column style map.
	 */
	private Map<Integer, CellStyle> columnStyleMap;

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