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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Generator {
	private static final String MODELE_RAPPORT = "com/multiplexSimoaGenerator/neuro4plex_Model.xlsx";
	private static final String PATH_DATA_OUTPUT = "C:/Data";
	private Map<String, List<ExcelRow>> beadPlexMap = new HashMap();

	public void execute() throws IOException {
		readInputFileAndStoreObjects();
	}

	private void readInputFileAndStoreObjects() throws IOException {
		FileInputStream file = new FileInputStream(new File("C:/Temp/input data.xlsx"));
		Workbook workbook = new XSSFWorkbook(file);
		Sheet sheet = workbook.getSheetAt(0);

		Map<Integer, List<String>> data = new HashMap();
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
						System.out.println("no AEB on line: " + (i + 1) + ", error message: "
								+ SheetUtil.getCellStringValue(row.getCell(SheetUtil.ERROR_TXT)));
					}
				}
			} catch (Exception e) {
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

		System.out.println("Total Number of BeadPlex found: " + beadPlexMap.keySet().size());
	}

	public static void buildExcel() {
		LocalDateTime timePoint = LocalDateTime.now();
		DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss");

		String formatDateTime = timePoint.format(formatter);

		InputStream is = Generator.class.getClassLoader()
				.getResourceAsStream("com/multiplexSimoaGenerator/neuro4plex_Model.xlsx");
		String fileName = "C:/Data/neuro4plex_" + formatDateTime + ".xlsx";
		File rapportOut = new File(fileName);
		copyFile(is, rapportOut);

		InputStream inputStream = null;
		try {
			inputStream = new FileInputStream(fileName);
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
			return;
		}

		FileOutputStream fos = null;
		
		try {
			fos = new FileOutputStream(fileName);
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
			size = -1;
			e.printStackTrace();
		}
		return size;
	}
}
