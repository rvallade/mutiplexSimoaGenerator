package com.multiplexSimoaGenerator;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class SheetUtil {
	public enum ColumnEnum {
		STATUS, BEAD_PLEX_NAME, SAMPLE_ID, TYPE, LOCATION, AEB, CONCENTRATION, FITTED_CONCENTRATION, ERRORS;
	}

	public static final IndexedColors[] colors = {IndexedColors.LEMON_CHIFFON, IndexedColors.CORAL, IndexedColors.TAN, IndexedColors.LIGHT_CORNFLOWER_BLUE, IndexedColors.YELLOW};

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