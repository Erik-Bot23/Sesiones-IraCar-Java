package com.erikjarquin.agendadesesiones.Service;

import com.erikjarquin.agendadesesiones.DTO.ExcelSelectColumnsResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.text.Normalizer;
import java.util.*;
import java.util.stream.Collectors;

@Service
public class ExcelSelectService {

    public ExcelSelectColumnsResponse parseSelectingColumns(MultipartFile file,
                                                            Integer sheetIndex,
                                                            List<String> desiredColumns,
                                                            int headerSearchRows,
                                                            int stopAfterEmptyRows) throws Exception {
        if (file == null || file.isEmpty()) {
            throw new IllegalArgumentException("No se recibió archivo.");
        }
        if (sheetIndex == null || sheetIndex < 0) sheetIndex = 0;
        if (headerSearchRows <= 0) headerSearchRows = 50;
        if (stopAfterEmptyRows <= 0) stopAfterEmptyRows = 20;

        try (InputStream is = file.getInputStream();
             Workbook wb = WorkbookFactory.create(is)) {

            if (sheetIndex >= wb.getNumberOfSheets()) {
                throw new IllegalArgumentException("Índice de hoja inválido. El archivo tiene " + wb.getNumberOfSheets() + " hojas.");
            }

            Sheet sheet = wb.getSheetAt(sheetIndex);
            String sheetName = wb.getSheetName(sheetIndex);
            DataFormatter fmt = new DataFormatter();

            // Normalizamos nombres objetivo (sin acentos/mayúsculas)
            List<String> targets = desiredColumns.stream()
                    .map(ExcelSelectService::normalize)
                    .collect(Collectors.toList());

            // Buscamos la fila de encabezados (la primera donde aparezcan TODOS los targets)
            HeaderDetection hd = findHeaderRowAndColumns(sheet, fmt, targets, headerSearchRows);
            if (hd == null || hd.colIndexByTarget.size() < targets.size()) {
                throw new IllegalArgumentException("No se pudieron localizar todas las columnas objetivo: " + desiredColumns);
            }

            int dataStart = hd.headerRow + 1;

            // Leemos filas desde dataStart, extrayendo solo los índices detectados
            List<List<String>> rows = new ArrayList<>();
            int emptyInARow = 0;
            for (int r = dataStart; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) { 
                    emptyInARow++;
                    if (emptyInARow >= stopAfterEmptyRows) break;
                    else continue;
                }

                List<String> out = new ArrayList<>();
                boolean allEmpty = true;
                for (String t : targets) {
                    int c = hd.colIndexByTarget.get(t);
                    String val = getCellAsString(fmt, row, c);
                    if (val != null && !val.isBlank()) allEmpty = false;
                    out.add(val);
                }

                if (allEmpty) {
                    emptyInARow++;
                    if (emptyInARow >= stopAfterEmptyRows) break;
                } else {
                    emptyInARow = 0;
                    rows.add(out);
                }
            }

            // Devuelve las columnas en su forma “bonita” original (las que pidió el usuario)
            return new ExcelSelectColumnsResponse(
                    sheetName,
                    desiredColumns,
                    rows,
                    hd.headerRow,
                    dataStart
            );
        }
    }

    // === Utilidades ===

    private static class HeaderDetection {
        int headerRow;
        Map<String, Integer> colIndexByTarget; // clave = target normalizado, valor = índice de columna
    }

    private static HeaderDetection findHeaderRowAndColumns(Sheet sheet, DataFormatter fmt, List<String> targets, int searchRows) {
        int lastRow = Math.min(sheet.getLastRowNum(), searchRows);
        int maxCols = estimateMaxCols(sheet, 5);

        for (int r = 0; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Map<String, Integer> found = new HashMap<>();
            for (int c = 0; c < maxCols; c++) {
                // Obtenemos texto del posible encabezado en esta celda,
                // resolviendo combinaciones/múltiples filas arriba
                String headerTxt = headerTextWithMergedFallback(sheet, fmt, r, c, 2).trim();
                if (headerTxt.isBlank()) continue;

                String norm = normalize(headerTxt);
                for (String target : targets) {
                    if (norm.equals(target)) {
                        found.putIfAbsent(target, c);
                    }
                }
            }

            if (found.size() == targets.size()) {
                HeaderDetection hd = new HeaderDetection();
                hd.headerRow = r;
                hd.colIndexByTarget = found;
                return hd;
            }
        }
        return null;
    }

    private static int estimateMaxCols(Sheet sheet, int sampleRows) {
        int max = 0;
        int limit = Math.min(sheet.getLastRowNum(), sampleRows);
        for (int r = 0; r <= limit; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > max) {
                max = row.getLastCellNum();
            }
        }
        if (max <= 0) max = 50; // fallback
        return max;
    }

    private static String headerTextWithMergedFallback(Sheet sheet, DataFormatter fmt, int rowIndex, int colIndex, int lookBehindRows) {
        // 1) Texto directo
        String txt = getCellTextResolvingMerged(sheet, fmt, rowIndex, colIndex);
        if (!txt.isBlank()) return txt;

        // 2) Si viene vacío (porque encabezado está arriba), mira hacia arriba N filas
        for (int i = 1; i <= lookBehindRows; i++) {
            int r = rowIndex - i;
            if (r < 0) break;
            String t2 = getCellTextResolvingMerged(sheet, fmt, r, colIndex);
            if (!t2.isBlank()) return t2;
        }
        return "";
    }

    private static String getCellTextResolvingMerged(Sheet sheet, DataFormatter fmt, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) return "";
        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        String txt = cell != null ? fmt.formatCellValue(cell) : "";
        if (txt != null && !txt.isBlank()) return txt;

        // Si está dentro de una región combinada, toma el valor de la celda superior izquierda de la región
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, colIndex)) {
                Row topRow = sheet.getRow(region.getFirstRow());
                if (topRow == null) break;
                Cell topCell = topRow.getCell(region.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                String mergedTxt = topCell != null ? fmt.formatCellValue(topCell) : "";
                return mergedTxt == null ? "" : mergedTxt;
            }
        }
        return "";
    }

    private static String getCellAsString(DataFormatter fmt, Row row, int colIndex) {
        Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        String v = cell == null ? "" : fmt.formatCellValue(cell);
        return v == null ? "" : v;
    }

    private static String normalize(String s) {
        if (s == null) return "";
        String n = Normalizer.normalize(s, Normalizer.Form.NFD)
                .replaceAll("\\p{M}+", "");            // quita acentos
        return n.toLowerCase().replaceAll("\\s+", " ").trim();
    }
}

