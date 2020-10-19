package util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

public class CheckCellNull {

	public static boolean isCellEmpty(final Cell cell) {
		if (cell == null) {
			return true;
		}

		if (cell.getCellType() == CellType.BLANK) {
			return true;
		}

		if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().trim().isEmpty()) {
			return true;
		}

		return false;
	}

}
