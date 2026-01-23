package com.erikjarquin.agendadesesiones.Controller;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.erikjarquin.agendadesesiones.DTO.ExcelParseRsponse;
import com.erikjarquin.agendadesesiones.Service.ExcelService;
import com.erikjarquin.agendadesesiones.Service.ExcelSelectService;

import jakarta.validation.constraints.Min;

@RestController
@RequestMapping("/api/excel")
@Validated
public class ExcelController {
    //private final ExcelService excelService;
    //public ExcelController(ExcelService excelService){
    //    this.excelService = excelService; 
    //}

    //@PostMapping(value = "/parse", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    //public ResponseEntity<ExcelParseRsponse> parseExcel(
    //    @RequestPart("file") MultipartFile file,
    //    @RequestParam(name = "sheetIndex", required = false) @Min(0) Integer sheetIndex
    //) throws Exception {
    //    ExcelParseRsponse response = excelService.parseExcel(file, sheetIndex);
    //    return ResponseEntity.ok(response);
    //}

    private final ExcelSelectService excelSelectService;
    public ExcelController(ExcelSelectService excelSelectService){
        this.excelSelectService = excelSelectService; 
    }

    @PostMapping(value = "/parse/select", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> parseExcelSelectingColumns(
        @RequestPart("file") MultipartFile file,
        @RequestParam(name="sheetIndex", required = false) Integer sheetIndex,
        @RequestParam(name = "colums", required = false, defaultValue = "UBICACIÓN, NOMBRE, FECHA, HORA") String columnsCsv,
        @RequestParam(name = "headerSearchRows", required = false, defaultValue = "60") Integer headerSearchRows,
        @RequestParam(name = "stopAfterEmptyRows", required = false, defaultValue = "20") Integer stopAfterEmptyRows
    ) throws Exception {
        //Permite columnas separadas por coma, con/ sin espacios
        List<String> desired = Arrays.stream(columnsCsv.split(",")).map(String::trim).filter(s -> !s.isBlank()).collect(Collectors.toList());
        var resp = excelSelectService.parseSelectingColumns(file, sheetIndex, desired, headerSearchRows, stopAfterEmptyRows);
        return ResponseEntity.ok(resp);
    }

    @GetMapping("/ping")
    public String ping(){
        return "Microservicio OK";
    }
}
