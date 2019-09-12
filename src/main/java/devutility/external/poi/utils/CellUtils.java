package devutility.external.poi.utils;

import java.math.BigDecimal;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;

import devutility.internal.data.converter.ConverterUtils;

public class CellUtils {
	/**
	 * Get cell's value.
	 * @param cell Cell object.
	 * @return Object
	 */
	public static Object getValue(Cell cell) {
		if (cell == null) {
			return null;
		}

		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();

		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				return cell.getDateCellValue();
			}

			return cell.getNumericCellValue();

		case BOOLEAN:
			return cell.getBooleanCellValue();

		case FORMULA:
			return cell.getCellFormula();

		case BLANK:
			return null;

		default:
			return cell;
		}
	}

	/**
	 * Get cell's value.
	 * @param cell Cell object.
	 * @param clazz Class object for field.
	 * @return {@code T}
	 */
	public static <T> T getValue(Cell cell, Class<T> clazz) {
		Object value = CellUtils.getValue(cell);

		if (value == null || value instanceof Cell) {
			return null;
		}

		String str = String.valueOf(value);

		if (value instanceof Double) {
			Double doubleValue = (Double) value;
			BigDecimal bigDecimalValue = new BigDecimal(doubleValue);
			str = bigDecimalValue.toPlainString();
		}

		return ConverterUtils.stringToType(str, clazz);
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
		CellType cellType = templateCell.getCellType();

		Cell cell = row.createCell(cellNum);
		cell.setCellComment(templateCell.getCellComment());
		cell.setCellStyle(templateCell.getCellStyle());

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