package devutility.external.poi.model;

/**
 * 
 * ExcelProperties
 * 
 * @author: Aldwin Su
 * @version: 2019-12-04 23:00:53
 */
public class ExcelProperties {
	private int maxSheetsCount;
	private int maxRowsCount;
	private int maxColumnsCount;

	public ExcelProperties() {
	}

	public ExcelProperties(int maxSheetsCount, int maxRowsCount, int maxColumnsCount) {
		this.maxSheetsCount = maxSheetsCount;
		this.maxRowsCount = maxRowsCount;
		this.maxColumnsCount = maxColumnsCount;
	}

	public int getMaxSheetsCount() {
		return maxSheetsCount;
	}

	public void setMaxSheetsCount(int maxSheetsCount) {
		this.maxSheetsCount = maxSheetsCount;
	}

	public int getMaxRowsCount() {
		return maxRowsCount;
	}

	public void setMaxRowsCount(int maxRowsCount) {
		this.maxRowsCount = maxRowsCount;
	}

	public int getMaxColumnsCount() {
		return maxColumnsCount;
	}

	public void setMaxColumnsCount(int maxColumnsCount) {
		this.maxColumnsCount = maxColumnsCount;
	}
}