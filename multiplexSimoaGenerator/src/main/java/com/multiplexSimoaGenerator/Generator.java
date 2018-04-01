package com.multiplexSimoaGenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {
	private static final String MODELE_RAPPORT = "com/multiplexSimoaGenerator/neuro4plex_Model.xlsx";
	private static final String PATH_DATA_OUTPUT = "C:/multiplexSimoaGenerator";
	private Map<String, List<ExcelRow>> beadPlexMap = new HashMap();
	private List<ExcelRow> errorRows = new ArrayList<ExcelRow>();

	public void execute() throws IOException {
		readInputFileAndStoreObjects();
	}
	
	private void readInputFileAndStoreObjects() throws IOException {
		Map<Integer, List<String>> data = new HashMap();
		FileInputStream inputfile = new FileInputStream(new File("C:/multiplexSimoaGenerator/input_data.xlsx"));
		
		// create empty result file
		String filename = getFilename();
		FileInputStream outputStream = buildExcel(filename);
		
		// read input file and build the beadPlex map
		buildBeadPlexMapFromInputFile(inputfile);
		System.out.println("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());
		
		// based on the map, create the tabs and fill them
		// for each beadPlex ==> 1 tab
		XSSFWorkbook wb = new XSSFWorkbook(outputStream);
		for (String key : beadPlexMap.keySet()) {
			XSSFSheet sheet = wb.cloneSheet(0, key);

			/*for (int r = 0; r < 5; r++) {
				XSSFRow row = sheet.createRow(r);

				for (int c = 0; c < 5; c++) {
					XSSFCell cell = row.createCell(c);

					cell.setCellValue("Cell " + r + " " + c);
				}
			}*/
		}
		wb.removeSheetAt(0);
		wb.setSheetOrder("ERRORS", wb.getNumberOfSheets()-1);
		wb.setActiveSheet(0);
		// add errors
		XSSFSheet errorSheet = wb.getSheet("ERRORS");
		int i = 1;
		for (ExcelRow currentRow : errorRows) {
			XSSFRow row = errorSheet.createRow(i++);
			XSSFCell cell = row.createCell(0);
			cell.setCellValue(currentRow.getErrorMessage());
		}
		FileOutputStream fileOut = new FileOutputStream(filename);

		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		
	}
	
	private void buildBeadPlexMapFromInputFile(FileInputStream file) throws IOException {
		Workbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		int i = 0;
		for (Row row : sheet) {
			try {
				if (i != 0) {
					String beadPlex = SheetUtil.getCellStringValue(row.getCell(SheetUtil.BEAD_PLEX_NAME));
					String aeb = SheetUtil.getCellStringValue(row.getCell(SheetUtil.AEB));

					Location location = new Location(SheetUtil.getCellStringValue(row.getCell(SheetUtil.LOCATION)));
					ExcelRow currentRow = new ExcelRow(
							i + 1, 
							beadPlex, 
							SheetUtil.getCellStringValue(row.getCell(SheetUtil.SAMPLE_ID)),
							SheetUtil.getCellStringValue(row.getCell(SheetUtil.CONCENTRATION)), 
							location, 
							aeb);
					
					if (!StringUtil.isEmpty(aeb)) {
						List<ExcelRow> rowsForBeadPlex = beadPlexMap.get(beadPlex);
						
						if (rowsForBeadPlex == null) {
							rowsForBeadPlex = new ArrayList();
							beadPlexMap.put(beadPlex, rowsForBeadPlex);
						}

						rowsForBeadPlex.add(currentRow);

						System.out.println(currentRow.toString());
					} else {
						currentRow.setErrorMessage(SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
						errorRows.add(currentRow);
						System.out.println("no AEB on line: " + (i + 1) + ", error message: "
								+ SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
					}
				}
			} catch (Exception e) {
				String beadPlex = SheetUtil.getCellStringValue(row.getCell(SheetUtil.BEAD_PLEX_NAME));
				String aeb = SheetUtil.getCellStringValue(row.getCell(SheetUtil.AEB));

				Location location = new Location(SheetUtil.getCellStringValue(row.getCell(SheetUtil.LOCATION)));
				ExcelRow currentRow = new ExcelRow(
						i + 1, 
						beadPlex, 
						SheetUtil.getCellStringValue(row.getCell(SheetUtil.SAMPLE_ID)),
						SheetUtil.getCellStringValue(row.getCell(SheetUtil.CONCENTRATION)), 
						location, 
						aeb);
				errorRows.add(currentRow);
				System.out.println("Error on line: " + (i + 1));
				System.out.println(e);
			}
			i++;
		}
		
		// sort the rows in the map for each key
		for (String key : beadPlexMap.keySet()) {
			Collections.sort(beadPlexMap.get(key), new Comparator<ExcelRow>() {

				public int compare(ExcelRow o1, ExcelRow o2) {
					Location loc1 = o1.getLocation();
					Location loc2 = o2.getLocation();
					if (loc1.getLetter().equals(loc2.getLetter())) {
						return loc1.getLetter().compareTo(loc2.getLetter());
					} else {
						return loc1.getNumber().compareTo(loc2.getNumber());
					}
				}
				
			});
		}
	}

	public static FileInputStream buildExcel(String filename) throws FileNotFoundException {
		InputStream is = Generator.class.getClassLoader().getResourceAsStream(MODELE_RAPPORT);
		File rapportOut = new File(filename);
		copyFile(is, rapportOut);

		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(filename);
		} catch (FileNotFoundException e) {
			System.out.println("File not found in the specified path.");
			e.printStackTrace();
		}

		XSSFSheet dataSheet;
		XSSFWorkbook workBook;
		try {
			workBook = new XSSFWorkbook(inputStream);
			dataSheet = workBook.getSheet("DATAS");
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}

		FileOutputStream fos = null;
		
		try {
			fos = new FileOutputStream(filename);
			workBook.write(fos);
		} catch (IOException e) {
			e.printStackTrace();

			if (fos != null) {
				try {
					fos.flush();
					fos.close();
				} catch (IOException ioe) {
					ioe.printStackTrace();
				}
			}
		} finally {
			if (fos != null) {
				try {
					fos.flush();
					fos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		
		return new FileInputStream(new File(filename));
	}

	private static int copyFile(InputStream input, File dest) {
		int size = -1;
		try {
			/*
			 * if (!src.exists()) throw new Exception("File doesn't exist: " +
			 * src.getPath());
			 */
			if (dest.exists())
				dest.delete();
			File dir = new File(dest.getParent());
			if (!dir.exists())
				dir.mkdirs();

			// FileInputStream input = new FileInputStream(src);
			FileOutputStream out = new FileOutputStream(dest);
			size = input.available();
			byte[] data = new byte[4096];
			int len;
			while ((len = input.read(data)) != -1) {
				out.write(data, 0, len);
			}
			out.flush();
			out.close();
			input.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return size;
	}
	
	private String getFilename() {
		LocalDateTime timePoint = LocalDateTime.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");
		String formatDateTime = timePoint.format(formatter);
		
		return "C:/multiplexSimoaGenerator/neuro4plex_" + formatDateTime + ".xlsx";
	}


}
