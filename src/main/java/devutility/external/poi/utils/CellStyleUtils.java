package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class CellStyleUtils {
	/**
	 * Clone CellStyle by provided Workbook and CellStyle.
	 * @param workbook: Workbook object.
	 * @param cellStyle: CellStyle object.
	 * @return CellStyle
	 */
	public static CellStyle clone(Workbook workbook, CellStyle cellStyle) {
		CellStyle newCellStyle = workbook.createCellStyle();
		newCellStyle.setAlignment(cellStyle.getAlignmentEnum());
		newCellStyle.setBorderBottom(cellStyle.getBorderBottomEnum());
		newCellStyle.setBorderLeft(cellStyle.getBorderLeftEnum());
		newCellStyle.setBorderRight(cellStyle.getBorderRightEnum());
		newCellStyle.setBorderTop(cellStyle.getBorderTopEnum());
		newCellStyle.setBottomBorderColor(cellStyle.getBottomBorderColor());
		newCellStyle.setDataFormat(cellStyle.getDataFormat());
		newCellStyle.setFillBackgroundColor(cellStyle.getFillBackgroundColor());
		newCellStyle.setFillForegroundColor(cellStyle.getFillForegroundColor());
		newCellStyle.setFillPattern(cellStyle.getFillPatternEnum());
		newCellStyle.setHidden(cellStyle.getHidden());
		newCellStyle.setIndention(cellStyle.getIndention());
		newCellStyle.setLeftBorderColor(cellStyle.getLeftBorderColor());
		newCellStyle.setLocked(cellStyle.getLocked());
		newCellStyle.setQuotePrefixed(cellStyle.getQuotePrefixed());
		newCellStyle.setRightBorderColor(cellStyle.getRightBorderColor());
		newCellStyle.setRotation(cellStyle.getRotation());
		newCellStyle.setShrinkToFit(cellStyle.getShrinkToFit());
		newCellStyle.setTopBorderColor(cellStyle.getTopBorderColor());
		newCellStyle.setVerticalAlignment(cellStyle.getVerticalAlignmentEnum());
		newCellStyle.setWrapText(cellStyle.getWrapText());
		return newCellStyle;
	}
}