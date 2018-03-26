package com.multiplexSimoaGenerator;

import org.apache.poi.ss.usermodel.Cell;

public class SheetUtil {
	public static final int SAMPLE_ID = 1;
	public static final int LOCATION = 3;
	public static final int BEAD_PLEX_NAME = 7;
	public static final int STATUS = 9;
	public static final int AEB = 12;
	public static final int CONCENTRATION = 13;
	public static final int ERROR_TXT = 33;

	public SheetUtil() {
	}

	public static String getCellStringValue(Cell cell) {
		if (cell == null)
			return "";
		if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.STRING.getCode())
			return cell.getStringCellValue();
		if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC.getCode()) {
			return Double.toString(cell.getNumericCellValue());
		}

		return "";
	}
}