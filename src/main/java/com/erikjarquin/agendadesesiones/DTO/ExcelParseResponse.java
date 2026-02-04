package com.erikjarquin.agendadesesiones.DTO;
import java.util.List;

public record ExcelParseResponse(String sheetName, List<String> headers, List<List<String>> rows, int rowcount) {
    public ExcelParseResponse(String sheetName, List<String> headers, List<List<String>> rows){
        this(
            sheetName, 
            headers != null ? headers : List.of(), 
            rows != null ? rows : List.of(), 
            rows != null ? rows.size() : 0
        );
    }
}
