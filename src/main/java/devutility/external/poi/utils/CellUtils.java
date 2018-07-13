package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;

public class CellUtils {
	/**
	 * Get cell's value.
	 * @param cell: Cell object.
	 * @param dataFormatter: DataFormatter object.
	 * @return Object
	 */
	public static Object getValue(Cell cell, DataFormatter dataFormatter) {
		if (cell == null) {
			return null;
		}

		switch (cell.getCellTypeEnum()) {
		case STRING:
			return cell.getRichStringCellValue().getString();

		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			} else {
				return cell.getNumericCellValue();
			}

		case BOOLEAN:
			return cell.getBooleanCellValue();

		case FORMULA:
			return cell.getCellFormula();

		case BLANK:
			return null;

		default:
			return dataFormatter.formatCellValue(cell);
		}
	}

	/**
	 * Clone a cell by template, add it into row.
	 * @param templateCell:Template Cell object.
	 * @param row: Row object.
	 * @return Cell
	 */
	public static Cell clone(Cell templateCell, Row row) {
		if (templateCell == null) {
			return null;
		}

		int cellNum = templateCell.getColumnIndex();
		CellType cellType = templateCell.getCellTypeEnum();

		Cell cell = row.createCell(cellNum);
		cell.setCellComment(templateCell.getCellComment());
		cell.setCellStyle(templateCell.getCellStyle());
		cell.setCellType(cellType);

		switch (cellType) {
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(templateCell)) {
				cell.setCellValue(templateCell.getDateCellValue());
				break;
			}

			cell.setCellValue(templateCell.getNumericCellValue());
			break;

		case STRING:
			String value = templateCell.getStringCellValue();

			if (value != null) {
				cell.setCellValue(templateCell.getStringCellValue());
				break;
			}

			RichTextString richTextString = templateCell.getRichStringCellValue();

			if (richTextString != null) {
				cell.setCellValue(richTextString);
			}

			break;

		case FORMULA:
			cell.setCellFormula(templateCell.getCellFormula());
			break;

		case BOOLEAN:
			cell.setCellValue(templateCell.getBooleanCellValue());
			break;

		case ERROR:
			cell.setCellErrorValue(templateCell.getErrorCellValue());
			break;

		default:
			break;
		}

		return cell;
	}
}