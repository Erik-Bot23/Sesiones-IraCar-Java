package com.erikjarquin.agendadesesiones.Service;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

//IMPORTS DE DE EXCEL
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//DTO DE SESIONES
import com.erikjarquin.agendadesesiones.DTO.ExcelParseResponse;

@Service
public class ExcelService {
    
// Límites defensivos (ajusta según tu caso)
    private static final int MAX_ROWS = 100_000;   // para evitar respuestas gigantes
    private static final int MAX_COLS = 200;

    public ExcelParseResponse parseExcel(MultipartFile file, Integer sheetIndex) throws IOException {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("No se recibió archivo");
        }

        if (sheetIndex == null || sheetIndex < 0) {
            sheetIndex = 0;
        }

        try (InputStream is = file.getInputStream();
             Workbook workbook = WorkbookFactory.create(is)) {
            if (sheetIndex >= workbook.getNumberOfSheets()) {
                throw new IllegalArgumentException("Índice de hoja inválido. El archivo tiene " 
                        + workbook.getNumberOfSheets() + " hojas.");
            }

            Sheet sheet = workbook.getSheetAt(sheetIndex);
            String sheetName = workbook.getSheetName(sheetIndex);

            // Usa Locale MX para formatos (fechas, números, etc.)
            DataFormatter formatter = new DataFormatter(Locale.forLanguageTag("es-MX"));

            // 1) Detectar la fila de encabezados (primera fila no vacía). Si quieres siempre la 0, omite este bloque.
            int headerRowIdx = findFirstNonEmptyRow(sheet, formatter);
            if (headerRowIdx < 0) {
                // Hoja vacía
                return new ExcelParseResponse(sheetName, List.of(), List.of());
            }

            // 2) Leer datos desde headerRowIdx hasta el final
            List<List<String>> data = new ArrayList<>();
            int lastRow = Math.min(sheet.getLastRowNum(), headerRowIdx + MAX_ROWS);
            for (int r = headerRowIdx; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                List<String> cells = new ArrayList<>();

                if (row != null) {
                    // Determina lastCell de forma segura y limitada
                    int lastCell = row.getLastCellNum();
                    if (lastCell < 0) lastCell = 0;
                    lastCell = Math.min(lastCell, MAX_COLS);

                    for (int c = 0; c < lastCell; c++) {
                        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        String value = (cell == null) ? "" : formatter.formatCellValue(cell);
                        cells.add(value == null ? "" : value.trim());
                    }
                    // Recortar trailing blanks
                    trimRightEmpty(cells);
                }
                data.add(cells);
            }

            // 3) Separar headers y rows
            List<String> headers = data.isEmpty() ? List.of() : data.get(0);
            List<List<String>> rows;
            if (data.size() > 1) {
                rows = data.subList(1, data.size());
            } else {
                rows = List.of();
            }
            return new ExcelParseResponse(sheetName, headers, rows);
        }
    }
    //Encuentra la primera fila con al menos una celda no vacía (formateada)
    private static int findFirstNonEmptyRow(Sheet sheet, DataFormatter formatter){
        int lastRow = sheet.getLastRowNum();
        for (int r = 0; r <= lastRow; r++){
            Row row = sheet.getRow(r);
            if (row == null) continue;

            int lastCell=row.getLastCellNum();
            if (lastCell < 0) continue;

            for(int c = 0; c < Math.min(lastCell, MAX_COLS); c++){
                Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) continue;
                String v = formatter.formatCellValue(cell);
                if (v != null && !v.trim().isEmpty()){
                    return r;
                }
            }
        }
        return -1;
    }

    //Elimina columnas vacís al final de la fila (["a","b","",""] -> ["a","b"])
    private static void trimRightEmpty(List<String> cells){
        int i = cells.size() - 1;
        while (i >= 0 && (cells.get(i) == null || cells.get(i).isBlank())) {
            i--;
        }
        if (i < cells.size() - 1){
            cells.subList(i+1, cells.size()).clear();
        }
    }
}


