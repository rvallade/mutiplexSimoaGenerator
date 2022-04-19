package com.multiplexSimoaGenerator;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {
	private Map<String, Integer> mapPositions = null;
	private static final String MODELE_RAPPORT = "com/multiplexSimoaGenerator/neuro4plex_Model.xlsx";
	private Map<String, BeadPlexBean> beadPlexMap = null;
	private List<String[]> rawData = null;
	private int nbRowsInSrcFile = 0;
	private int nbExcelRowProcessed = 0;
	private List<String> stringFromSrcFile = new ArrayList<>();
	private boolean sampleNameUsedAsIsInDuplicate = false;
	
	public void execute() throws IOException {
		log("Multiplex Simoa Generator - V2.2");
		log("START");
		File dir = new File("C:/multiplexSimoaGenerator");
		File[] files = dir.listFiles((d, name) -> name.endsWith(".csv"));
		
		for (File srcFile : files) {
			clearAttributes();
			
			FileInputStream inputfile = new FileInputStream(srcFile);
			// create empty result file
			String filename = getFilename(srcFile.getName());
			FileInputStream outputStream = buildExcel(filename);
			XSSFWorkbook wb = new XSSFWorkbook(outputStream);
			
			// read input file and build the beadPlex map
			try {
				log("Processing file (" + filename + ")");
				buildBeadPlexMapFromInputFile(inputfile);
				
				log("Read input file ... 100%");
				
				log("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());

				/*for (String key : beadPlexMap.keySet()) {
					log(beadPlexMap.get(key).toString());
				}*/
				
				log("Write output file ...");
				filloutNewFile(filename, wb);
				
				log("Write output file ... beadplex tabs 100%");
				filloutRowDataTab(wb);
				
				if (nbRowsInSrcFile != nbExcelRowProcessed) {
					log("######## WARNING: some rows missing in result file(" + nbRowsInSrcFile + "," + nbExcelRowProcessed + ").");
				}

				if (!stringFromSrcFile.isEmpty()) {
					for (String string : stringFromSrcFile) {
						log(string);
					}
				}
				log("Write output file ... raw data tab 100%");
				
				log("Reorder sheets and set active sheet...");
				wb.setSheetOrder("ERRORS", wb.getNumberOfSheets() - 1);
				wb.setSheetOrder("RAW DATA", wb.getNumberOfSheets() - 1);
				wb.setActiveSheet(0);
				
				log("Reorder sheets and set active sheet... 100%");
			} catch (Exception e) {
				log("An error occured, process stopped. You will find the root cause in the ERRORS tab.");
				e.printStackTrace();
				logErrorInExcelFile(e.getMessage(), wb);
				log("Loging error... 100%");
			}

			FileOutputStream fileOut = new FileOutputStream(filename);

			wb.setForceFormulaRecalculation(true);
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
		}
		log("END");
	}
	
	private void clearAttributes() {
		mapPositions = new HashMap<String, Integer>();
		beadPlexMap = new HashMap<String, BeadPlexBean>();
		rawData = new ArrayList<>();
		stringFromSrcFile = new ArrayList<>();
		nbExcelRowProcessed = 0;
		nbRowsInSrcFile = 0;
	}

	private void filloutRowDataTab(XSSFWorkbook wb) {
		XSSFSheet sheet = wb.getSheet("RAW DATA");
		int i = 0;
		for (String[] data : rawData) {
			XSSFRow row = sheet.createRow(i++);
			int j = 0;
			for (String value : data) {
				XSSFCell cell = getCell(row, j++);
				cell.setCellValue(value.replace("\"", ""));
			}
		}
	}

	private void filloutNewFile(String filename, XSSFWorkbook wb) throws IOException, FileNotFoundException {
		// based on the map, create the tabs and fill them
		// for each beadPlex ==> 1 tab
	    int color = 0;
		for (String key : beadPlexMap.keySet()) {
			XSSFSheet sheet = wb.cloneSheet(0, key);
		    IndexedColors currentColor = SheetUtil.colors[color++];
		    sheet.setTabColor(new XSSFColor(currentColor));
			try {
				fillSheet(sheet, key);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		
		wb.removeSheetAt(0);
	}
	
	private void logErrorInExcelFile(String message, XSSFWorkbook wb) {
		XSSFSheet errorSheet = wb.getSheet("ERRORS");
		XSSFRow row = errorSheet.createRow(1);
		XSSFCell cell = row.createCell(0);
		cell.setCellValue(message);
	}
 	
	private void fillSheet(XSSFSheet sheet, String key) throws Exception {
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
		
		//log(beadPlexBean.toString());
		
		int currentRow = 1;
		boolean twoRows = false;
		// CAL first
		currentRow = fillSheetForCalAndQC(beadPlexBean.getCalRows(), sheet, currentRow);
		// QC
		currentRow = fillSheetForCalAndQC(beadPlexBean.getQcRows(), sheet, currentRow);
		// OTHER ROWS
		Map<Integer, List<ExcelRow>> mapToProcess = beadPlexBean.getMapPositionExcelRows();
		for (int j = 1 ; j<50 ; j=j+2) {
			List<ExcelRow> list = mapToProcess.get(j);
			List<ExcelRow> duplicatesList = beadPlexBean.getDuplicateRows();
			
			if (list != null) {
				for (int i = 0 ; i < list.size() ; i++) {
					ExcelRow excelRow = list.get(i);
					//log("Processing main line: \r\n" + excelRow.toString());
					if (duplicatesList == null) {
						log("#######   duplicatesList is null");
					}
					// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
					ExcelRow potentialDuplicate = duplicatesList != null ? getDuplicateRow(duplicatesList, excelRow.getSampleID()) : null;
					if (potentialDuplicate != null) {
						//log("Processing duplicate: \r\n" + potentialDuplicate.toString());
						twoRows = true;
					}
					
					// the first is always there
					XSSFRow row = sheet.getRow(currentRow);
					
					getCell(row, 1).setCellValue(StringUtil.getCommonSampleName(excelRow.getSampleID(), sampleNameUsedAsIsInDuplicate));
					getCell(row, 2).setCellValue(excelRow.getLocation().toString());
					if (StringUtil.isEmpty(excelRow.getBeadPlex())) {
						getCell(row, 5).setCellValue(excelRow.getErrorMessage());
					} else {
						if (!StringUtil.isEmpty(excelRow.getAeb())) {
							getCell(row, 5).setCellValue(Double.parseDouble(excelRow.getAeb()));
						} else {
							getCell(row, 5).setCellValue(excelRow.getErrorMessage());
						}
						if (!StringUtil.isEmpty(excelRow.getFittedConcentration())) {
							getCell(row, 10).setCellValue(Double.parseDouble(excelRow.getFittedConcentration()));
						}
					}
					nbExcelRowProcessed++;
					stringFromSrcFile.remove(excelRow.getBeadPlex() + "/" + excelRow.getSampleID() + "/" + excelRow.getLocation().toString());
					
					if (twoRows) {
						getCell(row, 3).setCellValue(potentialDuplicate.getLocation().toString());
						if (StringUtil.isEmpty(potentialDuplicate.getBeadPlex())) {
							getCell(row, 6).setCellValue(potentialDuplicate.getErrorMessage());
						} else {
							if (!StringUtil.isEmpty(potentialDuplicate.getAeb())) {
								getCell(row, 6).setCellValue(Double.parseDouble(potentialDuplicate.getAeb()));
							} else {
								getCell(row, 6).setCellValue(potentialDuplicate.getErrorMessage());
							}
							if (!StringUtil.isEmpty(potentialDuplicate.getFittedConcentration())) {
								getCell(row, 11).setCellValue(Double.parseDouble(potentialDuplicate.getFittedConcentration()));
							}
						}
						twoRows = false;
						nbExcelRowProcessed++;
						stringFromSrcFile.remove(potentialDuplicate.getBeadPlex() + "/" + potentialDuplicate.getSampleID() + "/" + potentialDuplicate.getLocation().toString());
					}
					currentRow++;
				}
			}
		}
		
		//@TODO delete unused rows
		for (int i = currentRow ; i < 101 ; i++) {
			XSSFRow row = sheet.getRow(i);
			for (int j = 0 ; j < 14 ; j++) {
				getCell(row, j).setCellType(CellType.STRING);
				getCell(row, j).setCellValue("");
			}
		}
	}
	
	private void buildBeadPlexMapFromInputFile(FileInputStream file) throws Exception {
		InputStreamReader ipsr = new InputStreamReader(file);
		BufferedReader br = new BufferedReader(ipsr);

		List<ExcelRow> rowsWithoutBeadPlexlist = new ArrayList<>();

		String line;
		int i = 0;
		int rowNumber = 0;
		int posSampleID = -1;
		int posLocation = -1;
		int posBeadPleaxName = -1;
		int posStatus = -1;
		int posAEB = -1;
		int posConcentration = -1;
		int posFittedConcentration = -1;
		int posError = -1;
		int posType = -1;
				
		while ((line=br.readLine()) != null){
			// first we need to make sure we don't have any "" on the line
			if (line.contains("\"")) {
				// sometimes a number is in "" with a comma to separate thousands, like
				// ...AP,Complete,21,1.962676,"1,056.00",-,1,...
				// if we split using the comma only it produces a bug
				String[] datas = line.split("\"");
				for (int position = 1 ; position < datas.length ; ) {
					line = line.replace("\"" + datas[position] + "\"", datas[position].replace(",", ""));
					position = position + 2;
				}
			}
			
			String[] datas = line.split(",");
			rawData.add(datas);
			if (i == 0) {
				// HEADER: we need to get the position of each header 
				for (int z = 0 ; z < datas.length ; z++){
					mapPositions.put(datas[z].replaceAll("\"", ""), i++);
				}
				
				if (mapPositions.isEmpty()) {
					throw new Exception("Empty map of headers.");
				}
				
				if ((mapPositions.get(SheetUtil.SAMPLE_ID_LBL) == null && mapPositions.get(SheetUtil.SAMPLE_ID_LBL_ALT) == null) || mapPositions.get(SheetUtil.LOCATION_LBL) == null 
						|| mapPositions.get(SheetUtil.BEAD_PLEX_NAME_LBL) == null || mapPositions.get(SheetUtil.STATUS_LBL) == null
						|| mapPositions.get(SheetUtil.AEB_LBL) == null || mapPositions.get(SheetUtil.CONCENTRATION_LBL) == null
						|| mapPositions.get(SheetUtil.FITTED_CONCENTRATION_LBL) == null || mapPositions.get(SheetUtil.ERROR_TXT_LBL) == null
						|| mapPositions.get(SheetUtil.TYPE) == null) {
					throw new Exception("Can't find the loction of every relevant headers.");
				}
				
				posSampleID = mapPositions.get(SheetUtil.SAMPLE_ID_LBL);
				if (posSampleID == -1) {
					// meaning this file is using a newer version where SAMPLE ID is replaced with Name
					posSampleID = mapPositions.get(SheetUtil.SAMPLE_ID_LBL_ALT);
				}
				posLocation = mapPositions.get(SheetUtil.LOCATION_LBL);
				posBeadPleaxName = mapPositions.get(SheetUtil.BEAD_PLEX_NAME_LBL);
				posStatus = mapPositions.get(SheetUtil.STATUS_LBL);
				posAEB = mapPositions.get(SheetUtil.AEB_LBL);
				posConcentration = mapPositions.get(SheetUtil.CONCENTRATION_LBL);
				posFittedConcentration = mapPositions.get(SheetUtil.FITTED_CONCENTRATION_LBL);
				posError = mapPositions.get(SheetUtil.ERROR_TXT_LBL);
				posType = mapPositions.get(SheetUtil.TYPE);

				if (posSampleID == -1 || posLocation == -1 || posBeadPleaxName == -1 || posStatus == -1 
						|| posAEB == -1 || posConcentration == -1 || posError == -1 || posType == -1) {
					throw new Exception("Impossible to determine the correct position of all the relevant data.");
				}
				
				rowNumber++;
			} else {
				nbRowsInSrcFile++;
				// the actual data
				String beadPlex = datas[posBeadPleaxName].replaceAll("\"", "");
				String aeb = datas[posAEB].replaceAll("\"", "");

				Location location = new Location(datas[posLocation].replaceAll("\"", ""));
				ExcelRow currentRow = new ExcelRow(
						rowNumber++, 
						beadPlex,
						datas[posSampleID].replaceAll("\"", ""),
						datas[posConcentration].replaceAll("\"", ""), 
						location, 
						aeb,
						datas[posFittedConcentration].replaceAll("\"", ""),
						datas[posType]);
				currentRow.setErrorMessage(datas[posError].replaceAll("\"", ""));
				
				if (StringUtil.isEmpty(beadPlex)) {
					// add row to the list of rows without beadPlex. Those rows should be added to every beadPlex' map at the end
					rowsWithoutBeadPlexlist.add(currentRow);
					//log("Row added to rowsWithoutBeadPlexlist: " + currentRow.toString());
				} else {
					BeadPlexBean beadPlexBean = beadPlexMap.get(beadPlex);

					if (beadPlexBean == null) {
						beadPlexBean = new BeadPlexBean(beadPlex, sampleNameUsedAsIsInDuplicate);
						beadPlexMap.put(beadPlex, beadPlexBean);
					}

					beadPlexBean.addToGenericList(currentRow);
				}
				stringFromSrcFile.add(beadPlex + "/" + currentRow.getSampleID() + "/" + currentRow.getLocation().toString());
			}
		}
		/*log("List of rows before sort: ");
		for (String key : beadPlexMap.keySet()) {
			log(beadPlexMap.get(key).toString());
		}*/
		
		// we need to dispatch all the rows in different lists for each beadPlex
		for (String key : beadPlexMap.keySet()) {
			beadPlexMap.get(key).dispatchRows();
		}
		
		// add the rows without beadPlex in the map for each key
		for (String key : beadPlexMap.keySet()) {
			beadPlexMap.get(key).addRowsWithoutExplicitBeadPlex(rowsWithoutBeadPlexlist);
		}

		/*log("List of rows after sort: ");
		for (String key : beadPlexMap.keySet()) {
			log(beadPlexMap.get(key).toString());
		}*/

	}

	public FileInputStream buildExcel(String filename) throws FileNotFoundException {
		InputStream is = Generator.class.getClassLoader().getResourceAsStream(MODELE_RAPPORT);
		File rapportOut = new File(filename);
		copyFile(is, rapportOut);

		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(filename);
		} catch (FileNotFoundException e) {
			log("File not found in the specified path.");
			e.printStackTrace();
		}

		XSSFWorkbook workBook;
		try {
			workBook = new XSSFWorkbook(inputStream);
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}

		FileOutputStream fos = null;
		
		try {
			fos = new FileOutputStream(filename);
			workBook.write(fos);
			workBook.close();
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
			if (dest.exists()) {
				dest.delete();
			}
			File dir = new File(dest.getParent());
			if (!dir.exists()) {
				dir.mkdirs();
			}
			
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
	
	private String getFilename(String fileName) {
		return "C:/multiplexSimoaGenerator/" + fileName.substring(0, fileName.lastIndexOf(".")) + "_RESULT.xlsx";
	}

	private int fillSheetForCalAndQC(List<ExcelRow> list, XSSFSheet sheet, int currentRow) {
		boolean twoRows = false;
		for (int i = 0 ; i < list.size() ; ) {
			ExcelRow excelRow = list.get(i);
			//log("Processing: " + excelRow.toString());
			// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
			ExcelRow potentialDuplicate = null;
			if (i + 1 < list.size()) {
				potentialDuplicate = list.get(i+1);
			}
			
			if (potentialDuplicate != null && sampleNameUsedAsIsInDuplicate == false) {
				sampleNameUsedAsIsInDuplicate = StringUtil.isSameSampleNameForBothMode(excelRow.getSampleID(), potentialDuplicate.getSampleID());
			}
			if (potentialDuplicate != null 
					&& StringUtil.isSameSample(excelRow.getSampleID(), potentialDuplicate.getSampleID())) {
				twoRows = true;
			}
			
			// the first is always there
			XSSFRow row = sheet.getRow(currentRow);
			if (excelRow.isQCRow() && !StringUtil.isEmpty(excelRow.getConcentration())) {
				getCell(row, 12).setCellValue(Double.parseDouble(excelRow.getConcentration()));
			}
			getCell(row, 1).setCellValue(excelRow.isCalRow() ? "" : excelRow.getSampleID());
			getCell(row, 2).setCellValue(excelRow.getLocation().toString());
			
			if (StringUtil.isEmpty(excelRow.getBeadPlex())) {
				getCell(row, 5).setCellValue(excelRow.getErrorMessage());
			} else {
				if (!StringUtil.isEmpty(excelRow.getAeb())) {
					getCell(row, 5).setCellValue(Double.parseDouble(excelRow.getAeb()));
				} else {
					getCell(row, 5).setCellValue("");
				}
				if (!excelRow.isCalRow() && !StringUtil.isEmpty(excelRow.getFittedConcentration())) {
					getCell(row, 10).setCellValue(Double.parseDouble(excelRow.getFittedConcentration()));
				}
				if (excelRow.isCalRow()) {
					getCell(row, 12).setCellType(CellType.STRING);
					getCell(row, 12).setCellValue("");
					getCell(row, 13).setCellType(CellType.STRING);
					getCell(row, 13).setCellValue("");
				}
			}
			
			i++;
			
			stringFromSrcFile.remove(excelRow.getBeadPlex() + "/" + excelRow.getSampleID() + "/" + excelRow.getLocation().toString());
			
			nbExcelRowProcessed++;

			if (twoRows) {
				if (potentialDuplicate.isCalRow() && !StringUtil.isEmpty(potentialDuplicate.getConcentration())) {
					getCell(row, 0).setCellValue(Double.parseDouble(potentialDuplicate.getConcentration()));
				}
				getCell(row, 3).setCellValue(potentialDuplicate.getLocation().toString());
				if (StringUtil.isEmpty(potentialDuplicate.getBeadPlex())) {
					getCell(row, 6).setCellValue(potentialDuplicate.getErrorMessage());
				} else {
					if (!StringUtil.isEmpty(potentialDuplicate.getAeb())) {
						getCell(row, 6).setCellValue(Double.parseDouble(potentialDuplicate.getAeb()));
					} else {
						getCell(row, 6).setCellValue("");
					}
					if (!excelRow.isCalRow() && !StringUtil.isEmpty(potentialDuplicate.getFittedConcentration())) {
						getCell(row, 11).setCellValue(Double.parseDouble(potentialDuplicate.getFittedConcentration()));
					}
					if (excelRow.isCalRow()) {
						getCell(row, 12).setCellType(CellType.STRING);
						getCell(row, 12).setCellValue("");
						getCell(row, 13).setCellType(CellType.STRING);
						getCell(row, 13).setCellValue("");
					}
				}
				i++;
				
				nbExcelRowProcessed++;
				
				stringFromSrcFile.remove(potentialDuplicate.getBeadPlex() + "/" + potentialDuplicate.getSampleID() + "/" + potentialDuplicate.getLocation().toString());
			}
			currentRow++;
		}
		
		return currentRow;
	}
	
	private ExcelRow getDuplicateRow(List<ExcelRow> list, String sampleID) {
		ExcelRow duplicate = null;
		if (list != null) {
			for (ExcelRow row : list) {
				//log("Trying to find if duplicates " + sampleID + "/" + row.getSampleID() + "   ===> " + StringUtil.isSameSample(sampleID, row.getSampleID()));
				if (StringUtil.isSameSample(sampleID, row.getSampleID())) {
					duplicate = row;
					break;
				}
			}
		}
		return duplicate;
	}
	
	private XSSFCell getCell(XSSFRow row, int number) {
		return row.getCell(number, MissingCellPolicy.CREATE_NULL_AS_BLANK);
	}
	
	private void log(String message) {
		System.out.println(message);
	}
}
