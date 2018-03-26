package com.multiplexSimoaGenerator;

public class ExcelRow {
  private int id;
  private String beadPlex;
  private String sampleID;
  private String concentration;
  private Location location;
  private String aeb;
  private String errorMessage;
  
  public ExcelRow(int id, String beadPlex, String sampleID, String concentration, Location location, String aeb) {
    this.id = id;
    this.beadPlex = beadPlex;
    this.sampleID = sampleID;
    this.concentration = concentration;
    this.location = location;
    this.aeb = aeb;
  }
  
  public int getId() {
    return id;
  }
  
  public void setId(int id) { this.id = id; }
  
  public String getSampleID() {
    return sampleID;
  }
  
  public void setSampleID(String sampleID) { this.sampleID = sampleID; }
  
  public String getConcentration() {
    return concentration;
  }
  
  public void setConcentration(String concentration) { this.concentration = concentration; }
  
  public Location getLocation() {
    return location;
  }
  
  public void setLocation(Location location) { this.location = location; }
  
  public String getAeb() {
    return aeb;
  }
  
  public void setAeb(String aeb) { this.aeb = aeb; }
  
  public String getBeadPlex()
  {
    return beadPlex;
  }
  
  public void setBeadPlex(String beadPlex) {
    this.beadPlex = beadPlex;
  }
  
  public String getErrorMessage()
  {
    return errorMessage;
  }
  
  public void setErrorMessage(String errorMessage) {
    this.errorMessage = errorMessage;
  }
  
  public String toString()
  {
    return "Row #" + id + ", sampleID=" + sampleID + ", BeadPlex=" + beadPlex + ", location=" + location.toString() + ", Concentration=" + concentration + ", AEB=" + aeb;
  }
}