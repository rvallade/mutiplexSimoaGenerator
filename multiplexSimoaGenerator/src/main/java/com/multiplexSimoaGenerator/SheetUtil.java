package com.multiplexSimoaGenerator;

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
	
	public static final String SAMPLE_ID_LBL = "Sample ID";
	public static final String LOCATION_LBL = "Location";
	public static final String BEAD_PLEX_NAME_LBL = "Bead Plex Name";
	public static final String STATUS_LBL = "Status";
	public static final String AEB_LBL = "AEB";
	public static final String CONCENTRATION_LBL = "Conc.";
	public static final String FITTED_CONCENTRATION_LBL = "Fitted Conc"; 
	public static final String ERROR_TXT_LBL = "Errors";

	
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