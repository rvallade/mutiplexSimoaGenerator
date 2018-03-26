package com.multiplexSimoaGenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadWriteExcelFile3 {
	public ReadWriteExcelFile3() {
	}

	public static void readXLSXFile() throws IOException {
		InputStream ExcelFileToRead = new FileInputStream("C:/Temp/input data.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFWorkbook test = new XSSFWorkbook();

		XSSFSheet sheet = wb.getSheetAt(0);

		Iterator rows = sheet.rowIterator();

		while (rows.hasNext()) {
			XSSFRow row = (XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext()) {
				XSSFCell cell = (XSSFCell) cells.next();

				if (cell.getCellType() == 1) {
					System.out.print(cell.getStringCellValue() + " ");
				} else if (cell.getCellType() == 0) {
					System.out.print(cell.getNumericCellValue() + " ");
				}
			}

			System.out.println();
		}
	}

	public static void writeXLSXFile() throws IOException {
		String excelFileName = "C:/Temp/Test2.xlsx";

		String sheetName = "Sheet1";

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName);

		for (int r = 0; r < 5; r++) {
			XSSFRow row = sheet.createRow(r);

			for (int c = 0; c < 5; c++) {
				XSSFCell cell = row.createCell(c);

				cell.setCellValue("Cell " + r + " " + c);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(excelFileName);

		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}

	public static void main(String[] args) throws IOException {
		FileInputStream file = new FileInputStream(new File("C:/Temp/input data.xlsx"));
		Workbook workbook = new XSSFWorkbook(file);
		org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(0);

		Map<Integer, java.util.List<String>> data = new HashMap();
		int i = 0;
		try {
			for (Row row : sheet) {
				if (i != 0) {
					if (!getCellStringValue(row.getCell(9)).equals("Error")) {
						System.out.println(getCellStringValue(row.getCell(1)) + ";" + getCellStringValue(row.getCell(3))
								+ ";" + getCellStringValue(row.getCell(7)) + ";" + getCellStringValue(row.getCell(12))
								+ ";");
					} else {
						System.out.println("Error indicated on line: " + (i + 1));
					}
				}
				i++;
			}
		} catch (Exception e) {
			System.out.println("Error on line: " + (i + 1));
			System.out.println(e);
		}
	}

	private static String getCellStringValue(Cell cell) {
		if (cell.getCellType() == CellType.STRING.getCode()) {
			return cell.getStringCellValue();
		}
		if (cell.getCellType() == CellType.NUMERIC.getCode()) {
			return Double.toString(cell.getNumericCellValue());
		}

		return "";
	}
}