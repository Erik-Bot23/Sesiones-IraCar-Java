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
import com.erikjarquin.agendadesesiones.Service.ExcelSelectService;

@RestController
@RequestMapping("/api/excel")
@Validated
public class ExcelController {
    private final ExcelSelectService excelSelectService;
    public ExcelController(ExcelSelectService excelSelectService){
        this.excelSelectService = excelSelectService; 
    }

    @PostMapping(value = "/parse/select", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<?> parseExcelSelectingColumns(
        @RequestPart("file") MultipartFile file,
        @RequestParam(name="sheetIndex", required = false) Integer sheetIndex,
        // Columnas canónicas estandarizadas; las dem´sa se agregan solitas como dinámicas
        @RequestParam(name = "columns", required = false, defaultValue = "MODULO_UBICACIÓN, FECHA, HORA, TERMINAL_CF13") String columnsCsv,
        @RequestParam(name = "headerSearchRows", required = false, defaultValue = "60") Integer headerSearchRows,
        @RequestParam(name = "stopAfterEmptyRows", required = false, defaultValue = "20") Integer stopAfterEmptyRows,
        //Filtros canónicos
        @RequestParam(name = "terminalCf13", required = false) String filtroTerminalCf13,
        @RequestParam(name = "modulo", required = false) String filtroModuloUbic,
        @RequestParam(name = "fecha", required = false) String filtroFecha
    ) throws Exception {
        
        if (file == null || file.isEmpty()) {
                throw new IllegalArgumentException("No se recibió archivo");
            }
        if (sheetIndex == null || sheetIndex < 0) sheetIndex = 0;
        if (headerSearchRows == null || headerSearchRows < 1) headerSearchRows = 60;
        if (stopAfterEmptyRows == null || stopAfterEmptyRows < 1) stopAfterEmptyRows = 20;

        //Permite columnas separadas por coma, con/ sin espacios
        List<String> desired = Arrays.stream(columnsCsv.split(",")).map(String::trim).filter(s -> !s.isBlank()).collect(Collectors.toList());
        var resp = excelSelectService.parseSelectingColumns(file, sheetIndex, desired, headerSearchRows, stopAfterEmptyRows, filtroTerminalCf13, filtroModuloUbic, filtroFecha);
        return ResponseEntity.ok(resp);
    }

    @GetMapping("/ping")
    public String ping(){
        return "Microservicio OK";
    }
}
