package com.erikjarquin.agendadesesiones.DTO;

import java.util.List;

public class ExcelSelectColumnsResponse {
    private String sheetName;
    private List<String> columns; // Ubicación, nombre, fecha, hora
    private List<List<String>> rows;
    private int rowCount;
    private int headerRowIndex; // Fila donde se detectó el header (0-based)
    private int dataStarRowIndex; // Fila donde empiezan los datos (0-based)

    public ExcelSelectColumnsResponse(){}

    public ExcelSelectColumnsResponse(String sheetName, 
                                      List<String> columns, 
                                      List<List<String>> rows, 
                                      int headerRowIndex, 
                                      int dataStarRowIndex){
        this.sheetName = sheetName;
        this.columns = columns;
        this.rows = rows;
        this.rowCount = rows != null ? rows.size() : 0;
        this.headerRowIndex = headerRowIndex;
        this.dataStarRowIndex = dataStarRowIndex;
        }
    
    public String getSheetName(){ 
        return sheetName;
    }
    public void setSheetName(String sheetName){
        this.sheetName = sheetName;
    }

    public List<String> getColumns(){
        return columns;
    }
    public void setColumns(List<String> columns){
        this.columns = columns;
    }

    public List<List<String>> getRows(){
        return rows;
    }
    public void setRows(List<List<String>> rows){
        this.rows = rows;
    }

    public int getRowCount(){
        return rowCount;
    }
    public void setRowCount(int rowCount){
        this.rowCount = rowCount;
    }

    public int getHeaderRowIndex(){
        return headerRowIndex;
    }
    public void setHeaderRowIndex(int headerRowIndex){
        this.headerRowIndex = headerRowIndex;
    }

    public int getDataStartRowIndex(){
        return dataStarRowIndex;
    }
    public void setDataStartRowIndex(int dataStarRowIndex){
        this.dataStarRowIndex = dataStarRowIndex;
    }
}
