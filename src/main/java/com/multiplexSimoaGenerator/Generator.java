package com.multiplexSimoaGenerator;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.type.TypeFactory;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

public class Generator {
	private final String VERSION = "V4.0";
	private Map<String, Integer> mapPositions = null;
	private static final String MODEL_RAPPORT = "neuro4plex_Model.xlsx";
	private Map<String, BeadPlexBean> beadPlexMap = null;
	private List<String[]> rawData = null;
	private int nbRowsInSrcFile = 0;
	private int nbExcelRowProcessed = 0;
	private List<String> stringFromSrcFile = new ArrayList<>();
	private boolean sampleNameUsedAsIsInDuplicate = false;
	Logger logger = LogManager.getLogger(getClass());

	public void execute() throws Exception {
		log("\n" +
				"  __  __      _ _   _      _          ___ _                  ___                       _           \n" +
				" |  \\/  |_  _| | |_(_)_ __| |_____ __/ __(_)_ __  ___  __ _ / __|___ _ _  ___ _ _ __ _| |_ ___ _ _ \n" +
				" | |\\/| | || | |  _| | '_ \\ / -_) \\ /\\__ \\ | '  \\/ _ \\/ _` | (_ / -_) ' \\/ -_) '_/ _` |  _/ _ \\ '_|\n" +
				" |_|  |_|\\_,_|_|\\__|_| .__/_\\___/_\\_\\|___/_|_|_|_\\___/\\__,_|\\___\\___|_||_\\___|_| \\__,_|\\__\\___/_|  \n" +
				"                     |_|                                                                           \n");
		log("Multiplex Simoa Generator - " + VERSION);
		log("Provided by Romain Vallade (rvallade@gmail.com). Shoot a message if you like that program!");

		File dir = null;
		String os = System.getProperty("os.name");
		if (os.startsWith("Windows")) {
			System.out.println("Windows system");
			dir = new File("C:/multiplexSimoaGenerator");
		} else if (os.startsWith("Linux")) {
			System.out.println("Linux system");
			dir = new File(System.getProperty("user.home") + "/multiplexSimoaGenerator");
		}

		readConfigFileAndInitializeMapPositions(dir);

		File[] files = dir.listFiles((d, name) -> name.endsWith(".csv"));
		File[] resultFiles = dir.listFiles((d, name) -> name.endsWith(".xlsx"));

		if (files == null || files.length == 0) {
			log("0 file found.");
		} else {
			log(files.length + " files to process.");
			for (File srcFile : files) {
				clearAttributes();
				String srcFileName = srcFile.getName().substring(0, srcFile.getName().lastIndexOf("."));
				String filename = getFilename(srcFileName);
				FileInputStream inputFile = null;
				log("Processing file (" + srcFile.getName() + ")");

				// Do not continue if a result file already exists for that csv
				if (resultFiles != null && Arrays.stream(resultFiles).anyMatch(x -> x.getName().startsWith(srcFileName))) {
					log("\tA result file already exist. This file is skipped.");
					continue;
				}
				FileInputStream outputStream;
				XSSFWorkbook wb = null;

				// read input file and build the beadPlex map
				try {
					// create empty result file
					inputFile = new FileInputStream(srcFile);
					outputStream = buildExcel(filename);
					wb = new XSSFWorkbook(outputStream);

					buildBeadPlexMapFromInputFile(inputFile);

					log("Read input file ... 100%");

					log("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());

				/*for (String key : beadPlexMap.keySet()) {
					log(beadPlexMap.get(key).toString());
				}*/

					log("Write output file ...");
					filloutNewFile(wb);

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
					logger.error("An error occured, process stopped. You will find the root cause in the ERRORS tab.");
					e.printStackTrace();
					logErrorInExcelFile(e.getMessage(), wb);
					log("Logging error... 100%");
				}

				if (inputFile != null) {
					log("Save output file ...");
					FileOutputStream fileOut = new FileOutputStream(filename);

					wb.setForceFormulaRecalculation(true);
					wb.write(fileOut);
					fileOut.flush();
					fileOut.close();
				}
			}

		}

		log("END");
	}

	private void readConfigFileAndInitializeMapPositions(File dir) throws Exception {
		mapPositions = new HashMap<>();
		//read json file data to String
		List<ColumnJson> columnsMapping = null;
		try {
			byte[] jsonData = Files.readAllBytes(Paths.get(dir + "/columnMapping.txt"));
			//create ObjectMapper instance
			ObjectMapper objectMapper = new ObjectMapper();
			TypeFactory typeFactory = objectMapper.getTypeFactory();
			columnsMapping = objectMapper.readValue(jsonData, typeFactory.constructCollectionType(List.class, ColumnJson.class));
		} catch (Exception e) {

		}

		if (columnsMapping == null) {
			// if one doesn't exist, use the one in resources folder:
			log("A configuration file was not usable or none was provided, using the one from the jar.");
			try {
				InputStream is = Generator.class.getClassLoader().getResourceAsStream("columnMapping.txt");
				ObjectMapper objectMapper = new ObjectMapper();
				TypeFactory typeFactory = objectMapper.getTypeFactory();
				columnsMapping = objectMapper.readValue(is, typeFactory.constructCollectionType(List.class, ColumnJson.class));
			} catch (Exception e) {
				throw new Exception("Cannot find all the configuration file, ending process.");
			}
		}

		// set the map
		columnsMapping.forEach(
				x -> mapPositions.put(x.getKey(), x.getPosition() - 1)
		);

		// loop over all the objects in the enum to see if we have all the information we need
		if (Arrays.stream(SheetUtil.ColumnEnum.values()).anyMatch(x -> mapPositions.get(x.toString()) == null)) {
			// we are missing on of the columns in the config file, we cannot continue
			throw new Exception("Cannot find all the relevant columns in the configuration file.");
		}
	}

	private void clearAttributes() {
		beadPlexMap = new HashMap<>();
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

	private void filloutNewFile(XSSFWorkbook wb) {
		// based on the map, create the tabs and fill them
		// for each beadPlex ==> 1 tab
	    int color = 0;
		for (String key : beadPlexMap.keySet()) {
			XSSFSheet sheet = wb.cloneSheet(0, key);
		    IndexedColors currentColor = SheetUtil.colors[color++];
			byte[] rgb= DefaultIndexedColorMap.getDefaultRGB(currentColor.getIndex());
			sheet.setTabColor(new XSSFColor(rgb, null));
			try {
				fillSheet(wb, sheet, key);
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
 	
	private void fillSheet(Workbook workBook, XSSFSheet sheet, String key) {
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
		currentRow = fillSheetForCalAndQC(beadPlexBean.getCalRows(), sheet, currentRow, workBook);
		// QC
		currentRow = fillSheetForCalAndQC(beadPlexBean.getQcRows(), sheet, currentRow, workBook);
		// OTHER ROWS
		Map<Integer, List<ExcelRow>> mapToProcess = beadPlexBean.getMapPositionExcelRows();
		for (int j = 1 ; j < 50 ; j = j + 2) {
			List<ExcelRow> list = mapToProcess.get(j);
			List<ExcelRow> duplicatesList = beadPlexBean.getDuplicateRows();
			
			if (list != null) {
				for (ExcelRow excelRow : list) {
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
						getCell(row, 12).getCellStyle().setDataFormat(workBook.createDataFormat().getFormat("0.00"));
						getCell(row, 13).getCellStyle().setDataFormat(workBook.createDataFormat().getFormat("0.00"));
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
							getCell(row, 12).getCellStyle().setDataFormat(workBook.createDataFormat().getFormat("0.00"));
							getCell(row, 13).getCellStyle().setDataFormat(workBook.createDataFormat().getFormat("0.00"));
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
		int rowNumber = 0;
		int posSampleID = -1;
		int posLocation = -1;
		int posBeadPlexName = -1;
		int posStatus = -1;
		int posAEB = -1;
		int posConcentration = -1;
		int posFittedConcentration = -1;
		int posError = -1;
		int posType = -1;
				
		while ((line = br.readLine()) != null){
			// first we need to make sure we don't have any "" on the line
			if (line.contains("\"")) {
				// sometimes a number is in "" with a comma to separate thousands, like
				// ...AP,Complete,21,1.962676,"1,056.00",-,1,...
				// if we split using the comma only it produces a bug
				String[] data = line.split("\"");
				for (int position = 1 ; position < data.length ; ) {
					line = line.replace("\"" + data[position] + "\"", data[position].replace(",", ""));
					position = position + 2;
				}
			}
			
			String[] data = line.split(",");
			rawData.add(data);
			if (rowNumber == 0) {
				posStatus = mapPositions.get(SheetUtil.ColumnEnum.STATUS.toString());
				posBeadPlexName = mapPositions.get(SheetUtil.ColumnEnum.BEAD_PLEX_NAME.toString());
				posSampleID = mapPositions.get(SheetUtil.ColumnEnum.SAMPLE_ID.toString());
				posType = mapPositions.get(SheetUtil.ColumnEnum.TYPE.toString());
				posLocation = mapPositions.get(SheetUtil.ColumnEnum.LOCATION.toString());
				posAEB = mapPositions.get(SheetUtil.ColumnEnum.AEB.toString());
				posConcentration = mapPositions.get(SheetUtil.ColumnEnum.CONCENTRATION.toString());
				posFittedConcentration = mapPositions.get(SheetUtil.ColumnEnum.FITTED_CONCENTRATION.toString());
				posError = mapPositions.get(SheetUtil.ColumnEnum.ERRORS.toString());

				if (posSampleID == -1 || posLocation == -1 || posBeadPlexName == -1 || posStatus == -1
						|| posAEB == -1 || posConcentration == -1 || posError == -1 || posType == -1) {
					throw new Exception("Impossible to determine the correct position of all the relevant data.");
				}
				
				rowNumber++;
			} else {
				nbRowsInSrcFile++;
				// the actual data
				String beadPlex = data[posBeadPlexName].replaceAll("\"", "");
				String aeb = data[posAEB].replaceAll("\"", "");

				Location location = new Location(data[posLocation].replaceAll("\"", ""));
				ExcelRow currentRow = new ExcelRow(
						rowNumber++, 
						beadPlex,
						data[posSampleID].replaceAll("\"", ""),
						data[posConcentration].replaceAll("\"", ""),
						location, 
						aeb,
						data[posFittedConcentration].replaceAll("\"", ""),
						data[posType]);
				currentRow.setErrorMessage(data[posError].replaceAll("\"", ""));
				
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
		InputStream is = Generator.class.getClassLoader().getResourceAsStream(MODEL_RAPPORT);
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
		
		return new FileInputStream(filename);
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
		String path = null;
		String os = System.getProperty("os.name");
		if (os.startsWith("Windows")) {
			path = "C:/multiplexSimoaGenerator/";
		} else if (os.startsWith("Linux")) {
			path = System.getProperty("user.home") + "/multiplexSimoaGenerator/";
		}
		return path + fileName + "_RESULT-" + VERSION + ".xlsx";
	}

	private int fillSheetForCalAndQC(List<ExcelRow> list, XSSFSheet sheet, int currentRow, Workbook wb) {
		boolean twoRows = false;
		for (int i = 0 ; i < list.size() ; ) {
			ExcelRow excelRow = list.get(i);
			//log("Processing: " + excelRow.toString());
			// the next one should be the same sample, otherwise it means we have one of the 2 duplicates in error
			ExcelRow potentialDuplicate = null;
			if (i + 1 < list.size()) {
				potentialDuplicate = list.get(i+1);
			}
			
			if (potentialDuplicate != null && !sampleNameUsedAsIsInDuplicate) {
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
				getCell(row, 12).getCellStyle().setDataFormat(wb.createDataFormat().getFormat("0.00"));;
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
				} else {
					getCell(row, 13).getCellStyle().setDataFormat(wb.createDataFormat().getFormat("0.00"));;
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
						getCell(row, 13).getCellStyle().setDataFormat(wb.createDataFormat().getFormat("0.00"));
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
		logger.info(message);
	}
}
