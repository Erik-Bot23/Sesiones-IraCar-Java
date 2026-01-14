package com.erikjarquin.agendadesesiones.DTO;
import java.util.List;

public class ExcelParseRsponse {
    private String sheetName;
    private List<String> headers;
    private List<List<String>> rows;
    private int rowCount;

    public ExcelParseRsponse(){}
    public ExcelParseRsponse(String sheetName, List<String> headers, List<List<String>> rows){
        this.sheetName = sheetName;
        this.headers = headers;
        this.rows = rows;
        this.rowCount = rows != null ? rows.size() : 0;
    }

    public String getSheetName() { return sheetName; }
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<String> getHeaders() { return headers; }
    public void setHeaders(List<String> headers){
        this.headers = headers;
    }

    public List<List<String>> getRows(){ return rows; }
    public void setRows(List<List<String>> rows){
        this.rows = rows;
    }

    public int getRowCount(){ return rowCount; }
    public void setRowCount(int rowCount){
        this.rowCount = rowCount;
    }
    
}
