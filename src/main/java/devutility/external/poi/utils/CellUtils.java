package devutility.external.poi.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

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
}