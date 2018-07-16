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
		Map<Integer, List<String>> data = new HashMap();
		FileInputStream inputfile = new FileInputStream(new File("C:/multiplexSimoaGenerator/input_data.xlsx"));
		
		// create empty result file
		String filename = getFilename();
		FileInputStream outputStream = buildExcel(filename);
		
		// read input file and build the beadPlex map
		buildBeadPlexMapFromInputFile(inputfile);
		//System.out.println("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());
		
		// based on the map, create the tabs and fill them
		// for each beadPlex ==> 1 tab
		XSSFWorkbook wb = new XSSFWorkbook(outputStream);
		for (String key : beadPlexMap.keySet()) {
			XSSFSheet sheet = wb.cloneSheet(0, key);
			try {
				fillSheet(sheet, key);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
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

		wb.setForceFormulaRecalculation(true);
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
		
	}
	
	private void fillSheet(XSSFSheet sheet, String key) throws Exception {
		// just the common stuff
		XSSFRow header = sheet.getRow(0);
		header.getCell(0).setCellValue(key + " (pg/mL)");
		header.getCell(11).setCellValue("Final " + key + " (pg/mL)");
		
		List<ExcelRow> rowsForBeadPlex = beadPlexMap.get(key);
		Collections.sort(rowsForBeadPlex, new Comparator<ExcelRow>() {
			public int compare(ExcelRow o1, ExcelRow o2) {
				String sampleID1 = o1.getSampleID();
				String sampleID2 = o2.getSampleID();
				Location loc1 = o1.getLocation();
				Location loc2 = o2.getLocation();
				int n1 = Integer.parseInt(loc1.getNumber());
				int n2 = Integer.parseInt(loc2.getNumber());
				// CAL first, before anything else, then QC, then the rest
				if (sampleID1.toUpperCase().startsWith("CAL") || sampleID2.toUpperCase().startsWith("CAL")) {
					if (sampleID1.toUpperCase().startsWith("CAL") && sampleID2.toUpperCase().startsWith("CAL")) {
						return sampleID1.compareTo(sampleID2);
					} else if (sampleID1.toUpperCase().startsWith("CAL")){
						return -1;
					} else {
						return 1;
					}
				} else if (sampleID1.toUpperCase().startsWith("QC") || sampleID2.toUpperCase().startsWith("QC")) {
					if (sampleID1.toUpperCase().startsWith("QC") && sampleID2.toUpperCase().startsWith("QC")) {
						return sampleID1.compareTo(sampleID2);
					} else if (sampleID1.toUpperCase().startsWith("QC")){
						return -1;
					} else {
						return 1;
					}
				}  else if (!loc1.getNumber().equals(loc2.getNumber())) {
					return loc1.getNumber().compareTo(loc2.getNumber());
				} else {
					return loc1.getLetter().compareTo(loc2.getLetter());
				}
				
				/*if (n1 == n2) {
					return loc1.getLetter().compareTo(loc2.getLetter());
				} else {
					// same pair? ==> N1/N2 ==> N2 must be "even" and N1 must be N2-1
					// in a pair the biggest number must be evem
					if (n1 > n2) {
						if (n1 % 2 == 0 && n1 - 1 == n2) {
							// same pair
							return loc1.getLetter().compareTo(loc2.getLetter());
						} else {
							return loc1.getNumber().compareTo(loc2.getNumber());
						}
					} else {
						if (n2 % 2 == 0 && n2 - 1 == n1) {
							// same pair
							return loc1.getLetter().compareTo(loc2.getLetter());
						} else {
							return loc1.getNumber().compareTo(loc2.getNumber());
						}
					}
				}*/
			}
		});
		
		for (ExcelRow row : rowsForBeadPlex) {
			System.out.println(row);
		}
		
		int currentRow = 1;
		boolean twoRows = false;
		for (int i = 0 ; i <= rowsForBeadPlex.size() ; ) {
			ExcelRow excelRow = rowsForBeadPlex.get(i);
			//System.out.println("Processing: " + excelRow.toString());
			// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
			ExcelRow potentialDuplicate = null;
			if (i + 1 < rowsForBeadPlex.size()) {
				potentialDuplicate = rowsForBeadPlex.get(i+1);
			}
			if (potentialDuplicate != null) {
				if (StringUtil.isSameSample(excelRow.getSampleID(), potentialDuplicate.getSampleID())) {
					twoRows = true;
				}
			}
			
			// the first is always there
			XSSFRow row = sheet.getRow(currentRow);
			row.getCell(1).setCellValue(StringUtil.getSampleName(excelRow.getSampleID()));
			row.getCell(2).setCellValue(excelRow.getLocation().toString());
			if (StringUtil.isEmpty(excelRow.getBeadPlex())) {
				row.getCell(5).setCellValue(excelRow.getErrorMessage());
			} else {
				row.getCell(5).setCellValue(Double.parseDouble(excelRow.getAeb()));
			}
			i++;
			if (twoRows) {
				row.getCell(3).setCellValue(potentialDuplicate.getLocation().toString());
				if (StringUtil.isEmpty(potentialDuplicate.getBeadPlex())) {
					row.getCell(6).setCellValue(potentialDuplicate.getErrorMessage());
				} else {
					row.getCell(6).setCellValue(Double.parseDouble(potentialDuplicate.getAeb()));
				}
				i++;
			}
			currentRow++;
			// iterate on the list
			//for (int y = 0 ; y < rowsForBeadPlex.size() ; ) {
			//	ExcelRow excelRow = rowsForBeadPlex.get(y);
			//	if (Integer.parseInt(excelRow.getLocation().getNumber()) == i) {
			//		// first of the pair
			//		System.out.println(excelRow);
			//		System.out.println(beadPlexMap.get(key).get(y+1));
			//	}
			//	y = y + 2;
			//}
			//i = i + 2;
		}
		
		// delete unused rows
		for (int i = currentRow ; i < 102 ; i++) {
			SheetUtil.removeRow(sheet, i);
		}
	}
	
	private void buildBeadPlexMapFromInputFile(FileInputStream file) throws IOException {
		Workbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);
		List<ExcelRow> rowsWithoutBeadPlexlist = new ArrayList<ExcelRow>();

		int i = 0;
		for (Row row : sheet) {
			if (i != 0) {
				try {
					String beadPlex = SheetUtil.getCellStringValue(row.getCell(SheetUtil.BEAD_PLEX_NAME));
					String aeb = SheetUtil.getCellStringValue(row.getCell(SheetUtil.AEB));

					Location location = new Location(SheetUtil.getCellStringValue(row.getCell(SheetUtil.LOCATION)));
					ExcelRow currentRow = new ExcelRow(i + 1, 
							beadPlex,
							SheetUtil.getCellStringValue(row.getCell(SheetUtil.SAMPLE_ID)),
							SheetUtil.getCellStringValue(row.getCell(SheetUtil.CONCENTRATION)), 
							location, 
							aeb);
					currentRow.setErrorMessage(SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
					if (StringUtil.isEmpty(beadPlex)) {
						// add row to the list of rows without beadPlex. Those rows should be added to every beadPlex' map at the end
						rowsWithoutBeadPlexlist.add(currentRow);
					} else {
						if (!StringUtil.isEmpty(aeb)) {
							List<ExcelRow> rowsForBeadPlex = beadPlexMap.get(beadPlex);

							if (rowsForBeadPlex == null) {
								rowsForBeadPlex = new ArrayList<ExcelRow>();
								beadPlexMap.put(beadPlex, rowsForBeadPlex);
							}

							rowsForBeadPlex.add(currentRow);

							//System.out.println(currentRow.toString());
						} else {
							currentRow.setErrorMessage(SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
							errorRows.add(currentRow);
							//System.out.println("no AEB on line: " + (i + 1) + ", error message: " + SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
						}
					}
					
				} catch (Exception e) {
					ExcelRow exceptionRow = new ExcelRow(i + 1);
					exceptionRow.setIsException(true);
					exceptionRow.setErrorMessage("ALERT: unexpected issue with row " + exceptionRow.getId() + ", exception message: " + e.getMessage());
					errorRows.add(exceptionRow);
				}
			}
			i++;
		}

		// add the rows without beadPlex in the map for each key
		for (String key : beadPlexMap.keySet()) {
			beadPlexMap.get(key).addAll(rowsWithoutBeadPlexlist);
		}
		
		workbook.close();
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
