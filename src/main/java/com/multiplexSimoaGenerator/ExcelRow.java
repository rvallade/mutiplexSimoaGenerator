package com.multiplexSimoaGenerator;

public class ExcelRow {
	private int id;
	private String beadPlex;
	private String sampleID;
	private String concentration;
	private Location location;
	private String aeb;
	private String errorMessage;
	private String fittedConcentration;
	private boolean isException = false;
	private String type;

	public ExcelRow(int id, String beadPlex, String sampleID, String concentration, Location location, String aeb,
			String fittedConcentration, String type) {
		this.id = id;
		this.beadPlex = beadPlex;
		this.sampleID = sampleID;
		this.concentration = concentration;
		this.location = location;
		this.aeb = aeb;
		this.fittedConcentration = fittedConcentration;
		this.type = type;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public String getSampleID() {
		return sampleID;
	}

	public void setSampleID(String sampleID) {
		this.sampleID = sampleID;
	}

	public String getConcentration() {
		return concentration.replace("-", "");
	}

	public void setConcentration(String concentration) {
		this.concentration = concentration;
	}

	public Location getLocation() {
		return location;
	}

	public void setLocation(Location location) {
		this.location = location;
	}

	public String getAeb() {
		return aeb.replace("-", "");
	}

	public void setAeb(String aeb) {
		this.aeb = aeb;
	}

	public String getBeadPlex() {
		return beadPlex;
	}

	public void setBeadPlex(String beadPlex) {
		this.beadPlex = beadPlex;
	}

	public String getErrorMessage() {
		return errorMessage;
	}

	public void setErrorMessage(String errorMessage) {
		this.errorMessage = errorMessage;
	}

	public boolean isException() {
		return isException;
	}

	public void setIsException(boolean isException) {
		this.isException = isException;
	}

	public boolean isCalRow() {
		return "Calibrator".equalsIgnoreCase(type) || sampleID.toUpperCase().startsWith("CAL");
	}

	public boolean isQCRow() {
		return "Control".equalsIgnoreCase(type) || sampleID.toUpperCase().startsWith("QC");
	}

	public String getFittedConcentration() {
		return fittedConcentration.replace("-", "");
	}

	public void setFittedConcentration(String fittedConcentration) {
		this.fittedConcentration = fittedConcentration;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public String toString() {
		return "Row #" + id + "   \t sampleID=" + sampleID + "    \t BeadPlex=" + beadPlex + "       \t location="
				+ location.toString() + "\t Concentration=" + concentration + "\t AEB=" + aeb + "\t Fitted Conc.="
				+ fittedConcentration + "\t Type=" + type + "\t Error=" + errorMessage;
	}
}