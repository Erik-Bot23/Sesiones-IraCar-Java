package com.erikjarquin.agendadesesiones.DTO;

import java.util.List;

public record ExcelSelectColumnsResponse(String sheetName, List<String> columns, List<List<String>> rows, int rowCount, int headerRowIndex, int dataStartRowIndex) {
    public ExcelSelectColumnsResponse (String sheetName, List<String> columns, List<List<String>> rows, int headerRowIndex, int dataStartRowIndex){
        this(
            sheetName,
            columns != null ? columns : List.of(),
            rows != null ? rows : List.of(),
            rows != null ? rows.size() : 0,
            headerRowIndex,
            dataStartRowIndex
        );
    }
}
