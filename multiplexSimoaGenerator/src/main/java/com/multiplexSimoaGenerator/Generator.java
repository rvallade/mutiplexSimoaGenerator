package com.multiplexSimoaGenerator;

import java.awt.Font;
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

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {
	private static final String MODELE_RAPPORT = "com/multiplexSimoaGenerator/neuro4plex_Model.xlsx";
	private static final String PATH_DATA_OUTPUT = "C:/multiplexSimoaGenerator";
	private Map<String, BeadPlexBean> beadPlexMap = new HashMap<String, BeadPlexBean>();
	private List<ExcelRow> errorRows = new ArrayList<ExcelRow>();
	
	public void execute() throws IOException {
		Map<Integer, List<String>> data = new HashMap<Integer, List<String>>();
		FileInputStream inputfile = new FileInputStream(new File("C:/multiplexSimoaGenerator/input_data.xlsx"));
		
		// create empty result file
		String filename = getFilename();
		FileInputStream outputStream = buildExcel(filename);
		
		// read input file and build the beadPlex map
		buildBeadPlexMapFromInputFile(inputfile);
		System.out.println("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());
		for (String key : beadPlexMap.keySet()) {
			System.out.println(beadPlexMap.get(key).toString());
		}
		filloutNewFile(filename, outputStream);
		
	}

	private void filloutNewFile(String filename, FileInputStream outputStream) 
			throws IOException, FileNotFoundException {
		// based on the map, create the tabs and fill them
		// for each beadPlex ==> 1 tab
		XSSFWorkbook wb = new XSSFWorkbook(outputStream);

		XSSFCellStyle style1 = wb.createCellStyle();
	    style1.setFillForegroundColor(new XSSFColor(new java.awt.Color(128, 0, 128)));
	    style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    int color = 0;
	    
		for (String key : beadPlexMap.keySet()) {
			XSSFSheet sheet = wb.cloneSheet(0, key);
			XSSFCellStyle currentStyle = wb.createCellStyle();
		    IndexedColors currentColor = SheetUtil.colors[color++];
		    style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		    style1.setFillForegroundColor(new XSSFColor(currentColor));
		    style1.setVerticalAlignment(VerticalAlignment.CENTER);
		    style1.setAlignment(HorizontalAlignment.CENTER);
		    sheet.setTabColor(new XSSFColor(currentColor));
			try {
				fillSheet(sheet, key, currentColor);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
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
	
	private void fillSheet(XSSFSheet sheet, String key, IndexedColors currentColor) throws Exception {
		// just the common stuff
		XSSFRow header = sheet.getRow(0);
		header.getCell(0).setCellValue(key + " (pg/mL)");
		header.getCell(11).setCellValue("Final " + key + " (pg/mL)");
		/*for (int i = 0 ; i <= 11 ; i++) {
			header.getCell(i).getCellStyle().setFillPattern(FillPatternType.SOLID_FOREGROUND);
			header.getCell(i).getCellStyle().setFillForegroundColor(new XSSFColor(currentColor));
		}*/
		BeadPlexBean beadPlexBean = beadPlexMap.get(key);
		beadPlexBean.sortLists();
		
		//System.out.println(beadPlexBean.toString());
		
		int currentRow = 1;
		boolean twoRows = false;
		// CAL first
		currentRow = processList(beadPlexBean.getCalRows(), sheet, currentRow);
		// QC
		currentRow = processList(beadPlexBean.getQcRows(), sheet, currentRow);
		// OTHER ROWS
		Map<Integer, List<ExcelRow>> mapToProcess = beadPlexBean.getMapPositionExcelRows();
		for (int j = 1 ; j<50 ; j=j+2) {
			List<ExcelRow> list = mapToProcess.get(j);
			List<ExcelRow> duplicatesList = mapToProcess.get(j+1);
			
			if (list != null) {
				for (int i = 0 ; i < list.size() ; i++) {
					ExcelRow excelRow = list.get(i);
					System.out.println("Processing: " + excelRow.toString());
					// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
					ExcelRow potentialDuplicate = duplicatesList != null ? getDuplicateRow(duplicatesList, excelRow.getSampleID()) : null;
					twoRows = potentialDuplicate != null;
					
					// the first is always there
					XSSFRow row = sheet.getRow(currentRow);
					row.getCell(1).setCellValue(StringUtil.getCommonSampleName(excelRow.getSampleID(), potentialDuplicate != null ? potentialDuplicate.getSampleID() : null));
					row.getCell(2).setCellValue(excelRow.getLocation().toString());
					if (StringUtil.isEmpty(excelRow.getBeadPlex())) {
						row.getCell(5).setCellValue(excelRow.getErrorMessage());
					} else {
						row.getCell(5).setCellValue(Double.parseDouble(excelRow.getAeb()));
					}
					
					if (twoRows) {
						row.getCell(3).setCellValue(potentialDuplicate.getLocation().toString());
						if (StringUtil.isEmpty(potentialDuplicate.getBeadPlex())) {
							row.getCell(6).setCellValue(potentialDuplicate.getErrorMessage());
						} else {
							row.getCell(6).setCellValue(Double.parseDouble(potentialDuplicate.getAeb()));
						}
					}
					currentRow++;
				}
			}
		}
		
		// delete unused rows
		/*for (int i = currentRow ; i < 101 ; i++) {
			SheetUtil.removeRow(sheet, i);
		}*/
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
						BeadPlexBean beadPlexBean = beadPlexMap.get(beadPlex);

						if (beadPlexBean == null) {
							beadPlexBean = new BeadPlexBean(beadPlex);
							beadPlexMap.put(beadPlex, beadPlexBean);
						}

						beadPlexBean.addRow(currentRow);
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
			beadPlexMap.get(key).addRowsWithoutExplicitBeadPlex(rowsWithoutBeadPlexlist);
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

	private int processList(List<ExcelRow> list, XSSFSheet sheet, int currentRow) {
		boolean twoRows = false;
		for (int i = 0 ; i < list.size() ; ) {
			ExcelRow excelRow = list.get(i);
			//System.out.println("Processing: " + excelRow.toString());
			// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
			ExcelRow potentialDuplicate = null;
			if (i + 1 < list.size()) {
				potentialDuplicate = list.get(i+1);
			}
			if (potentialDuplicate != null) {
				if (StringUtil.isSameSample(excelRow.getSampleID(), potentialDuplicate.getSampleID())) {
					twoRows = true;
				}
			}
			
			// the first is always there
			XSSFRow row = sheet.getRow(currentRow);
			row.getCell(1).setCellValue(excelRow.isCalRow() ? "" : excelRow.getSampleID());
			row.getCell(2).setCellValue(excelRow.getLocation().toString());
			if (StringUtil.isEmpty(excelRow.getBeadPlex())) {
				row.getCell(5).setCellValue(excelRow.getErrorMessage());
			} else {
				if (!StringUtil.isEmpty(excelRow.getAeb())) {
					row.getCell(5).setCellValue(Double.parseDouble(excelRow.getAeb()));
				} else {
					row.getCell(5).setCellValue("");
				}
			}
			i++;
			if (twoRows) {
				row.getCell(3).setCellValue(potentialDuplicate.getLocation().toString());
				if (StringUtil.isEmpty(potentialDuplicate.getBeadPlex())) {
					row.getCell(6).setCellValue(potentialDuplicate.getErrorMessage());
				} else {
					if (!StringUtil.isEmpty(potentialDuplicate.getAeb())) {
						row.getCell(6).setCellValue(Double.parseDouble(potentialDuplicate.getAeb()));
					} else {
						row.getCell(6).setCellValue("");
					}
				}
				i++;
			}
			currentRow++;
		}
		
		return currentRow;
	}
	
	private ExcelRow getDuplicateRow(List<ExcelRow> list, String sampleID) {
		ExcelRow duplicate = null;
		if (list != null) {
			for (ExcelRow row : list) {
				if (StringUtil.isSameSample(sampleID, row.getSampleID())) {
					duplicate = row;
					break;
				}
			}
		}
		return duplicate;
	}
}
