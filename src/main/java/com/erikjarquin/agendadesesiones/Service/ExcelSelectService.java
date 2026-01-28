package com.erikjarquin.agendadesesiones.Service;

import com.erikjarquin.agendadesesiones.DTO.ExcelSelectColumnsResponse;

import org.apache.poi.ss.format.CellTextFormatter;
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
                                                            int stopAfterEmptyRows,
                                                            String filtroTerminal,
                                                            String filtroUbicación,
                                                            String filtroFecha) throws Exception {
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

            // Normaliza columnas deseadas (admite TERMINAL_CF13 / TERMINAL_CF12)
            List <String> targets = desiredColumns.stream().map(s -> {
                String n = normalize(s);
                if (n.equals("terminal_cf13") || n.equals("terminal_cf12")) return n;
                return n;
            }).toList();

            // Buscamos la fila de encabezados (la primera donde aparezcan TODOS los targets)
            HeaderDetection hd = findHeaderRowAndColumns(sheet, fmt, targets, headerSearchRows);
            if (hd == null || hd.colIndexByTarget.size() < targets.size()) {
                throw new IllegalArgumentException("No se pudieron localizar todas las columnas objetivo: " + desiredColumns);
            }

            int dataStart = hd.headerRow + 1;

            // Lee todas las filas
            List<List<String>> allRows = new ArrayList<>();
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
                    allRows.add(out);
                }
            }

            //=== Filtros ===//
            final List<String> headerPretty = desiredColumns; // Columnas como las pidió el usurio
            //Índices de columnas para filtrar (normalizados)
            int idxUbic = indexOfNormalized(headerPretty, "ubicación");

            // Esta línea para que sirve?, ya que si la desmarco me marco error en la variable idxUbic en el if de la línea 113
            //if (idxUbic < 0) idxUbic = indexOfNormalized(headerPretty, "ubicación"); 

            int idxFecha = indexOfNormalized(headerPretty, "fecha");

            // Índeces para terminales por centro (puede que existan varios, pero el filtro sÓlo se aplica a CF13)
            int idxTermCF13 = indexOfNormalized(headerPretty,"terminal_cf13");
            //Si además  muestras CF12/CF16/etc. Se verán en la tabla, pero no se usarán para filtrar
            int idxTermCF12 = indexOfNormalized(headerPretty,"terminal_cf12");
            int idxTermCF16 = indexOfNormalized(headerPretty,"terminal_cf16");
            // ... (puedes terner más si los agregas en columns)

            String fTerm = normalizeNullable(filtroTerminal);
            String fUbic = normalizeNullable(filtroUbicación);
            String fFecha = normalizeNullable(filtroFecha);

            List<List<String>> filtered = allRows.stream().filter(row ->{
                boolean ok = true;

                if (fUbic != null && idxUbic >= 0){
                    String v = normalize(row.get(idxUbic));
                    ok &= v.contains(fUbic);
                }

                if (fFecha != null && idxFecha >= 0){
                    String v = normalize(row.get(idxFecha));
                    ok &= v.contains(fFecha);
                }

                if (fTerm != null && idxTermCF13 >= 0){
                    String v13 = normalize(row.get(idxTermCF13));
                    ok &= v13.contains(fTerm);
                }
                return ok;
            }).toList();

            // Devuelve las columnas en su forma “bonita” original (las que pidió el usuario)
            return new ExcelSelectColumnsResponse(
                    sheetName,
                    headerPretty,
                    filtered,
                    hd.headerRow,
                    dataStart
            );
        }
    }

    //=== Helpers para filtros ===//
    private static int indexOfNormalized(List<String> headers, String target){
        String t = normalize(target);
        for (int i=0; i<headers.size();i++){
            String h = headers.get(i);
            if (normalize(h).equals(t)) return i;
        }
        return -1;
    }

    private static String normalizeNullable(String s){
        if (s==null || s.isBlank()) return null;
        return normalize(s);
    }

    // === Utilidades ===
    private static class HeaderDetection {
        int headerRow;
        Map<String, Integer> colIndexByTarget; // clave = target normalizado, valor = índice de columna
    }

    private static HeaderDetection findHeaderRowAndColumns(Sheet sheet, DataFormatter fmt, List<String> targetsRaw, int searchRows) {
        // Normaliza targets (ej. ubicación, fecha, hora, nombre, terminal_cf13, terminal_cf12, ...)
        List<String> targets = targetsRaw.stream().map(ExcelSelectService::normalize).toList();
        int lastRow = Math.min(sheet.getLastRowNum(), searchRows);
        int maxCols = estimateMaxCols(sheet, 5);

        for (int r=0; r<=lastRow; r++){
            Row row = sheet.getRow(r);
            if (row==null) continue;

            Map<String, Integer> found = new HashMap<>();
            for (int c=0; c<= maxCols; c++){
                String cellText = headerTextWithMergedFallback(sheet, fmt, r, c, 2).trim();
                if (cellText.isBlank()) continue;

                String normCell = normalize(cellText);
                String normPath = headerPathNormalized(sheet, fmt, r, c, 3); //Top->Down

                for(String target : targets){
                    if (!found.containsKey(target) && matchesTargetWithHierarchy(target, normCell, normPath)){
                        found.put(target, c);
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

    // Devuelve la "ruta" de encabezados: ejemplo "centro federal no.13 > datos de la cita > terminal"
    private static String headerPathNormalized(Sheet sheet, DataFormatter fmt, int rowIndex, int colIndex, int lookUpLevels){
        List<String> parts = new ArrayList<>();
        for (int i = 0; i <= lookUpLevels; i++){
            int r = rowIndex - i;
            if (r<0) break;
            String t = getCellTextResolvingMerged(sheet, fmt, rowIndex, colIndex);
            if (t != null && !t.isBlank()) {
                parts.add(normalize(t));
            } 
        }

        // La ruta de arriba hacia abajo (top->down) ayuda a "contiene...terminal"
        Collections.reverse(parts);
        return String.join(" > ", parts);
    }

    private static boolean matchesTargetWithHierarchy(String canonicalTarget, String normCell, String normPath){
       String target = normalizeCanonical(canonicalTarget);

        // 1) MODULO_UBICACION
        if (target.equals("modulo_ubicacion")) {
            //Coincide si la celda o cualquier segmento de la ruta está en los aliases
            if (MODULO_UBI_ALIASES.contains(normCell)) return true;
            for (String seg : normPath.split(">")){
                if (MODULO_UBI_ALIASES.contains(seg.trim())) return true;
            } 
            return true;
        }

        // 2) FECHA / HORA
        if (target.equals("fecha")) {
            if (FECHA_ALIASES.contains(normCell)) return true;
            for (String seg : normPath.split(">")) if(FECHA_ALIASES.contains(seg.trim())) return true;
            return false;
        }

        if (target.equals("hora")){
            if(HORA_ALIASES.contains(normCell)) return true;
            for (String seg : normPath.split(">")) if(HORA_ALIASES.contains(seg.trim())) return true;
            return false;
        }

        // 3) TERMINAL_CFXX (CF13, CF12, CF16, ...)
        if (target.startsWith("terminal_cf")) {
            String digits = extractCenterDigits(target); // "13"
            if (digits.isBlank()) return false;

            boolean hasTerminal = normPath.contains("terminal");
            // Segmento inferior
            // Nombres posibles del centro
            boolean cfFull = normPath.contains("centro federal no. "+ digits);
            boolean cpsSpaced = normPath.contains("cps " + digits) || normPath.contains("cefereso " + digits);
            boolean cpsCompact = normPath.contains("cps" + digits) || normPath.contains("cefereso" + digits);

            boolean centerMatch = cfFull || cpsSpaced || cpsCompact;
            return hasTerminal && centerMatch;
        }

        
    // 4) Otros (si agregas más campos canónicos)
    // Igualdad directa por celda o por cualquier segmento de ruta
    if (normCell.equals(target)) return true;
    for (String seg : normPath.split(">")) if (seg.trim().equals(target)) return true;

    return false;

    }

    // === ALIASES para nombres canónicos ===//
    //Se usará "MODULO_UBICACIÓN", "FECHA", "HORA", "TERMINAL_CF13/12/16...
    private static final Set<String> MODULO_UBI_ALIASES=Set.of("modulo", "módulo", "ubicacion", "ubicación", "ubicacion exp",
                                                                "ubicación exp", "modulo/ubicacion", "módulo/ubicación", "ubicacion.", "ubicación.");

    //En algunos libros FECHA/HORA vienen bajo "DATOS DE LA CITA"
    private static final Set<String> FECHA_ALIASES = Set.of("fecha", "f. cita", "fecha de cita");
    private static final Set<String> HORA_ALIASES =  Set.of("hora", "h. cita", "hora de cita");

    //Normaliza un label "canónico" (lo que el usuario pide 'columns')
    private static String normalizeCanonical(String s){
        return normalize(s).replace("/", "_").replace("-", "_").replace(".", "").trim();
    }

    // Convierte "TERMINAL_CF13" -> "13", "terminal_cf16" -> "16"
    private static String extractCenterDigits(String canonicalTarget){
        String norm = normalizeCanonical(canonicalTarget);
        if (!norm.startsWith("terminal_cf")) return "";
        return norm.replaceFirst("^terminal_cf\\s*", "");
    }
}

