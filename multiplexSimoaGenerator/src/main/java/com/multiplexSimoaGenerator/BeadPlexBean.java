package com.multiplexSimoaGenerator;

import java.util.List;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Map;

public class BeadPlexBean {
	private String beadPlex = null;
	private Map<Integer, List<ExcelRow>> mapPositionExcelRows = new HashMap<Integer, List<ExcelRow>>();
	private List<ExcelRow> calRows = new ArrayList<ExcelRow>();
	private List<ExcelRow> qcRows = new ArrayList<ExcelRow>();
	
	public BeadPlexBean(String beadPlex) {
		this.beadPlex = beadPlex;
	}
	
	public void addRow(ExcelRow row) {
		if (row.isCalRow()) {
			addToCalRows(row);
		} else if (row.isQCRow()) {
			addToQCRows(row);
		} else {
			addExcelRowToPositionMap(row);
		}
	}
	
	private void addExcelRowToPositionMap(ExcelRow row) {
		int position = Integer.parseInt(row.getLocation().getNumber());
		List<ExcelRow> list = mapPositionExcelRows.get(position);
		if (list == null) {
			list = new ArrayList<ExcelRow>();
			mapPositionExcelRows.put(position, list);
		}
		list.add(row);
	}
	
	private void addToCalRows(ExcelRow row) {
		calRows.add(row);
	}

	private void addToQCRows(ExcelRow row) {
		qcRows.add(row);
	}
	
	public void addRowsWithoutExplicitBeadPlex(List<ExcelRow> rows) {
		for(ExcelRow row : rows) {
			addExcelRowToPositionMap(row);
		}
	}
	
	public void sortLists() {
		Collections.sort(calRows, new Comparator<ExcelRow>() {
			public int compare(ExcelRow o1, ExcelRow o2) {
				String sampleID1 = o1.getSampleID();
				String sampleID2 = o2.getSampleID();
				if (sampleID1.toUpperCase().startsWith("CAL") && sampleID2.toUpperCase().startsWith("CAL")) {
					return sampleID1.compareTo(sampleID2);
				} else if (sampleID1.toUpperCase().startsWith("CAL")){
					return -1;
				} else {
					return 1;
				}
			}
		});
		Collections.sort(qcRows, new Comparator<ExcelRow>() {
			public int compare(ExcelRow o1, ExcelRow o2) {
				String sampleID1 = o1.getSampleID();
				String sampleID2 = o2.getSampleID();
				if (sampleID1.toUpperCase().startsWith("QC") && sampleID2.toUpperCase().startsWith("QC")) {
					return sampleID1.compareTo(sampleID2);
				} else if (sampleID1.toUpperCase().startsWith("QC")){
					return -1;
				} else {
					return 1;
				}
			}
		});
		for (int i = 1 ; i<50 ; i++) {
			List<ExcelRow> list = mapPositionExcelRows.get(i);
			if (list != null) {
				Collections.sort(list, new Comparator<ExcelRow>() {
					public int compare(ExcelRow o1, ExcelRow o2) {
						Location loc1 = o1.getLocation();
						Location loc2 = o2.getLocation();
						return loc1.getLetter().compareTo(loc2.getLetter());
					}
				});	
			}
		}
	}

	@Override
	public String toString() {
		StringBuilder stringBuilder = new StringBuilder();
		stringBuilder.append("\r\n ########### START\r\n");
		if (mapPositionExcelRows.isEmpty() && calRows.isEmpty() && qcRows.isEmpty()) {
			stringBuilder.append("All lists or maps are empty for ");
			stringBuilder.append(beadPlex);
		} else {
			stringBuilder.append("Lists or maps for beadPlex ");
			stringBuilder.append(beadPlex);
			stringBuilder.append("\r\n CAL:");
			for (ExcelRow row : calRows) {
				stringBuilder.append("\r\n");
				stringBuilder.append(row.toString());
			}
			stringBuilder.append("\r\n QC:");
			for (ExcelRow row : qcRows) {
				stringBuilder.append("\r\n");
				stringBuilder.append(row.toString());
			}
			stringBuilder.append("\r\n Other:");
			for (int i = 1 ; i<50 ; i++) {
				List<ExcelRow> list = mapPositionExcelRows.get(i);
				if (list != null) {
					stringBuilder.append("\r\n Position");
					stringBuilder.append(i);
					for (ExcelRow row : list) {
						stringBuilder.append("\r\n");
						stringBuilder.append(row.toString());
					}
				}
			}
		}
		stringBuilder.append("\r\n ########### END");		
		return stringBuilder.toString();
	}

	public Map<Integer, List<ExcelRow>> getMapPositionExcelRows() {
		return mapPositionExcelRows;
	}

	public List<ExcelRow> getCalRows() {
		return calRows;
	}

	public List<ExcelRow> getQcRows() {
		return qcRows;
	}
}
