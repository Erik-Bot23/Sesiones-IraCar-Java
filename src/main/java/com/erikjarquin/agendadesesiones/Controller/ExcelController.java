package com.erikjarquin.agendadesesiones.Controller;

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

import jakarta.validation.constraints.Min;

@RestController
@RequestMapping("/api/excel")
@Validated
public class ExcelController {
    private final ExcelService excelService;
    public ExcelController(ExcelService excelService){
        this.excelService = excelService; 
    }

    @PostMapping(value = "/parse", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<ExcelParseRsponse> parseExcel(
        @RequestPart("file") MultipartFile file,
        @RequestParam(name = "sheetIndex", required = false) @Min(0) Integer sheetIndex
    ) throws Exception {
        ExcelParseRsponse response = excelService.parseExcel(file, sheetIndex);
        return ResponseEntity.ok(response);
    }

    @GetMapping("/ping")
    public String ping(){
        return "Microservicio OK";
    }
}
