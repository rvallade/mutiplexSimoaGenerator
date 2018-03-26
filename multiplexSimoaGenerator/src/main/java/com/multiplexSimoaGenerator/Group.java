package com.multiplexSimoaGenerator;

import java.util.List;

public class Group
{
  private List<ExcelRow> rows = null;
  
  public Group() {
    rows = new java.util.ArrayList();
  }
  
  public List<ExcelRow> addRow(ExcelRow row) {
    rows.add(row);
    return rows;
  }
}