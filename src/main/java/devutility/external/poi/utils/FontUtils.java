package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class FontUtils {
	/**
	 * Get Font object from Workbook and CellStyle.
	 * @param workbook: Workbook object.
	 * @param cellStyle: CellStyle object.
	 * @return Font
	 */
	public static Font get(Workbook workbook, CellStyle cellStyle) {
		short index = cellStyle.getFontIndex();
		return workbook.getFontAt(index);
	}

	/**
	 * Get Font object from Workbook and Cell.
	 * @param workbook: Workbook object.
	 * @param cell: Cell object.
	 * @return Font
	 */
	public static Font get(Workbook workbook, Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return get(workbook, cellStyle);
	}

	/**
	 * Get Font object by Workbook and Row object.
	 * @param workbook: Workbook object.
	 * @param row: Row object.
	 * @return Font
	 * @throws IllegalAccessException
	 */
	public static Font get(Workbook workbook, Row row) throws IllegalAccessException {
		short cellNum = row.getFirstCellNum();

		if (cellNum == -1) {
			throw new IllegalAccessException("No cell found in row!");
		}

		Cell cell = row.getCell(cellNum);
		return get(workbook, cell);
	}

	/**
	 * Clone Font object by Workbook and Font object.
	 * @param workbook: Workbook object.
	 * @param font: Font object.
	 * @return Font
	 */
	public static Font clone(Workbook workbook, Font font) {
		Font newFont = workbook.createFont();
		newFont.setBold(font.getBold());
		newFont.setCharSet(font.getCharSet());
		newFont.setColor(font.getColor());
		newFont.setFontHeight(font.getFontHeight());
		newFont.setFontHeightInPoints(font.getFontHeightInPoints());
		newFont.setFontName(String.format("%s_copy", font.getFontName()));
		newFont.setItalic(font.getItalic());
		newFont.setStrikeout(font.getStrikeout());
		newFont.setTypeOffset(font.getTypeOffset());
		newFont.setUnderline(font.getUnderline());
		return newFont;
	}
}