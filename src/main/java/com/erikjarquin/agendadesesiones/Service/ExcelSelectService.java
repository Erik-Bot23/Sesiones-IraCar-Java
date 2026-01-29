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
                                                            List<String> desiredCanonicalColumns, //Ahora espera nombres canónicos
                                                            int headerSearchRows,
                                                            int stopAfterEmptyRows,
                                                            String filtroTerminalCf13,
                                                            String filtroModuloUbic,
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
            List <String> targets = desiredCanonicalColumns.stream().map(ExcelSelectService::normalizeCanonical).toList();

            // Buscamos la fila de encabezados (la primera donde aparezcan TODOS los targets)
            HeaderDetection hd = findHeaderRowAndColumns(sheet, fmt, targets, headerSearchRows);
            if (hd == null || hd.colIndexByTarget.size() < targets.size()) {
                throw new IllegalArgumentException("No se pudieron localizar todas las columnas objetivo: " + desiredCanonicalColumns);
            }

            int dataStart = hd.headerRow + 1;

            // Lee filas y arma valores normalizados por columna canónica
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

                for (String canonical : targets) {
                    int c = hd.colIndexByTarget.get(canonical); //Índice de esa columna detectada
                    String raw = getCellAsString(fmt, row, c);

                    String val = raw;
                    switch (canonical) {
                        case "modulo_ubicacion" -> {
                            val = (raw == null ? "" : raw.trim().toUpperCase());
                        }
                        case "fecha" -> val = normalizeDateValue(raw);
                        case "hora" -> val = normalizeTimeValue(raw);
                        default -> {
                            // Terminales u otros: dejar como están, solo trim
                            val = (raw == null ? "" : raw.trim()); 
                        }
                    }
                        if (!val.isBlank()) allEmpty = false;
                        out.add(val);
                }
                if (allEmpty){
                    if (++emptyInARow >= stopAfterEmptyRows) break;
                } else {
                    emptyInARow = 0;
                    allRows.add(out);
                }
            }

            //=== Filtros ===//
            final List<String> headerPretty = desiredCanonicalColumns; // Columnas como las pidió el usurio
            //Índices de columnas para filtrar (normalizados)
            int idxModulo = indexOfNormalized(headerPretty, "modulo_ubicación");
            int idxFecha = indexOfNormalized(headerPretty, "fecha");
            int idxTerm13 = indexOfNormalized(headerPretty, "terminal_cf13");

            String fModulo = normalizeNullable(filtroModuloUbic);
            String fFecha = normalizeNullable(filtroFecha);
            String fTerm13 = normalizeNullable(filtroTerminalCf13);

            List<List<String>> filtered = allRows.stream().filter(row ->{
                boolean ok = true;

                if (fModulo != null && idxModulo >= 0) ok &=normalize(row.get(idxModulo)).contains(fModulo);

                if (fFecha != null && idxFecha >= 0) ok &=normalize(row.get(idxFecha)).contains(fFecha);
                // SOLO aplica a CF13
                if (fTerm13 != null && idxTerm13 >= 0) ok &=normalize(row.get(idxTerm13)).contains(fTerm13);
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

    // Normaliza fecha a yyyy-MM-dd (acepta dd/MM/yyyy), dd-MM-yyyy, yyyy-MM-dd, ddMMyyyy)
    private static String normalizeDateValue(String s){
        if (s==null) return "";
        String v = s.trim().replace(".", "/").replace("-","/");
        try {
            // yyyy/MM/dd
            if (v.matches("\\d{4}/\\d{1,2}/\\d{1,2}")) {
                var dt = java.time.LocalDate.parse(v, java.time.format.DateTimeFormatter.ofPattern("yyyy/M/d"));
                return dt.toString(); // yyyy-MM-dd
            }
            // dd/MM/yyyy
            if (v.matches("\\d{1,2}/\\d{1,2}/\\d{4}")){
                var dt = java.time.LocalDate.parse(v,java.time.format.DateTimeFormatter.ofPattern("d/M/yyyy"));
                return dt.toString();
            }
            // ddMMyyyy
            if (v.matches("\\d{8}")){
                String dd = v.substring(0,2), mm = v.substring(2, 4), yyyy = v.substring(4, 8);
                var dt = java.time.LocalDate.parse(dd + "/" + mm + "/" + yyyy, java.time.format.DateTimeFormatter.ofPattern("d/M/yyyy"));
                return dt.toString();
            }
        } catch (Exception ignore){}
        return s.trim(); //fallback
    }

    // Normaliza hora a HH:mm (acepta H:mm, HH:mm, 11.00 -> 11:00)
    private static String normalizeTimeValue(String s){
        if (s == null) return "";
        String v = s.trim().replace(".", ":");
        try {
            if (v.matches("\\d{1,2}:\\{2}")){
                var t = java.time.LocalTime.parse(v, java.time.format.DateTimeFormatter.ofPattern("H:mm"));
                return t.toString().substring(0,5); //HH:mm
            }
        } catch (Exception ignore) {}
        return s.trim();
    }

    // Devuelve la ruta de encabezados en bruto (sin normalizar) top->down, separada por " > "
    private static String headerPathRaw(Sheet sheet, DataFormatter fmt, int rowIndex, int colIndex, int lookUpLevels){
        List<String> parts = new ArrayList<>();
        for (int i = 0; i <= lookUpLevels; i++){
            int r = rowIndex - i;
            if (r<0) break;
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIndex, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            String txt = cell != null ? fmt.formatCellValue(cell) : "";
            if (txt == null || txt.isBlank()){
                // si está en merged, toma la celda superior izquierda
                for (var region : sheet.getMergedRegions()){
                    if (region.isInRange(r, colIndex)) {
                        Row top = sheet.getRow(region.getFirstRow());
                        if (top!= null){
                            Cell tc = top.getCell(region.getFirstColumn(), Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                            txt = tc != null ? fmt.formatCellValue(tc) : "";
                        }
                        break;
                    }
                }
            }
            if (txt != null && !txt.isBlank()) parts.add(txt.trim());
        }
        Collections.reverse(parts);
        return String.join(" > ", parts);
    }

    //Representa una columna dinámica "Termina - {Centro}"
    private static class TerminalCol {
        String canonicalKey; // ej: terminal_cf16, terminal_marina, terminal_otro...
        String displayLabel; // ej: "Terminal - CEFERESO No. 16"
        int colIndex;
        TerminalCol(String key, String label, int idx){
            this.canonicalKey = key;
            this.displayLabel = label;
            this.colIndex = idx;
        }

        // Extrae (key, label) del path (normalizado y raw)
        private static TerminalCol extractTerminalKeyAndLabelFromPaths(String normPath, String rawPath, int colIndex){
            // Busca el segmento "centro" en el path crudo para el label
            String displayCenter = null;
            for(String seg : rawPath.split(">")){
                String s = seg.trim();
                String sn = normalize(s);
                if (sn.contains("centro federal no") || sn.contains("cefereso no.") || sn.matches(".*\\bcps\\b.*") || sn.contains("marina")){
                    displayCenter = s;
                    break;
                }
            }

            if (displayCenter == null){
                // Fallback: intenta deducir desde normpath
                displayCenter = "OTRO CENTRO";
            }

            // Detecta CFxx o MARINA para la key canónica
            String key = "terminal_otro";
            //CFxx por distintas formas  (centro federal / cefereso / cps)
            java.util.regex.Matcher m = java.util.regex.Pattern.compile(".*(?:centro federal no\\.|cefereso no\\.|cps)\\s*(\\d+).*").matcher(normPath);
            if (m.matches()){
                key = "terminal_cf" + m.group(1);
            } else if (normPath.contains("marina")) {
                key = "terminal_marina";
            }
            String label = "Terminal - " + displayCenter.toUpperCase();
            return new TerminalCol(key, label, colIndex);
        }
    }

    



    // === Utilidades ===
    private static class HeaderDetection {
        int headerRow;
        Map<String, Integer> colIndexByTarget; // clave = target normalizado, valor = índice de columna
    }

    private static HeaderDetection findHeaderRowAndColumns(Sheet sheet, DataFormatter fmt, List<String> canonicalTargets, int searchRows) {
       List<String> targets = canonicalTargets.stream().map(ExcelSelectService::normalizeCanonical).toList();
       int lastRow = Math.min(sheet.getLastRowNum(), searchRows);
       int maxCols = estimateMaxCols(sheet, 5);

       for (int r=0; r<=lastRow;r++){
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Map<String, Integer> found = new HashMap<>();
            for (int c = 0; c < maxCols; c++){
                String cellText = headerTextWithMergedFallback(sheet, fmt, r, c, 3).trim();
                if (cellText.isBlank()) continue;

                String normCell = normalize(cellText);
                String normPath = headerPathNormalized(sheet, fmt, r, c, 4); //top->down

                for (String target : targets){
                    if (!found.containsKey(target) && matchesTargetWithHierarchy(target, normCell, normPath)){
                        found.put(target, c);
                    }
                }
            }

            if (found.size() == targets.size()){
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

