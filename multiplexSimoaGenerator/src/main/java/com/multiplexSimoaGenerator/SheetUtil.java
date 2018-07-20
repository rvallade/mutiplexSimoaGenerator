package com.multiplexSimoaGenerator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class SheetUtil {
	public static final int SAMPLE_ID = 1;
	public static final int LOCATION = 3;
	public static final int BEAD_PLEX_NAME = 7;
	public static final int STATUS = 9;
	public static final int AEB = 12;
	public static final int CONCENTRATION = 13;
	public static final int ERROR_TXT = 33;
	
	public static final int OUT_CONCENTRATION = 0;
	public static final int OUT_SAMPLE_ID = 1;
	public static final int OUT_LOCATION_1 = 2;
	public static final int OUT_LOCATION_2 = 3;
	public static final int OUT_AEB_1 = 4;
	public static final int OUT_AEB_2 = 5;
	
	public static final IndexedColors[] colors = {IndexedColors.LEMON_CHIFFON, IndexedColors.CORAL, IndexedColors.TAN, IndexedColors.LIGHT_CORNFLOWER_BLUE, IndexedColors.YELLOW};
		
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
	
	public static void removeRow(Sheet sheet, int rowIndex) {
	    int lastRowNum = sheet.getLastRowNum();
	    if (rowIndex >= 0 && rowIndex < lastRowNum) {
	        sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
	    }
	    if (rowIndex == lastRowNum) {
	        Row removingRow = sheet.getRow(rowIndex);
	        if (removingRow != null) {
	            sheet.removeRow(removingRow);
	        }
	    }
	}
}