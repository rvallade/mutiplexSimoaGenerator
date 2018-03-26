package com.multiplexSimoaGenerator;

public class Location {
  private String letter;
  private String number;
  
  public Location(String letter, String number) {
    this.letter = letter;
    this.number = number;
  }
  
  public Location(String location) {
    buildLocationFromString(location);
  }
  
  private void buildLocationFromString(String location) {
    String[] split = location.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
    if (split.length == 2) {
      letter = split[0];
      number = split[1];
    }
  }
  
  public String getLetter() {
    return letter;
  }
  
  public void setLetter(String letter) { this.letter = letter; }
  
  public String getNumber() {
    return number;
  }
  
  public void setNumber(String number) { this.number = number; }
  

  public String toString()
  {
    return letter + number;
  }
}