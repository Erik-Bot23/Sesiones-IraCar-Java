package com.erikjarquin.agendadesesiones.Service;

import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;


//IMPORTS DE DE EXCEL
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

//DTO DE SESIONES
import com.erikjarquin.agendadesesiones.DTO.ExcelParseRsponse;

@Service
public class ExcelService {
    public ExcelParseRsponse parseExcel(MultipartFile file, Integer sheetIndex) throws Exception {
        if (file == null || file.isEmpty()){
            throw new IllegalArgumentException("No se recibió archivo");
        }

        if (sheetIndex == null || sheetIndex < 0){
            sheetIndex = 0;
        }

        try (InputStream is = file.getInputStream();
            Workbook workbook = WorkbookFactory.create(is)) {
                if (sheetIndex >= workbook.getNumberOfSheets()) {
                    throw new IllegalArgumentException("Índice de hoja invalido. El archivo tiene " + workbook.getNumberOfSheets() + "hojas.");
                }

                Sheet sheet = workbook.getSheetAt(sheetIndex);
                String sheetName = workbook.getSheetName(sheetIndex);

                DataFormatter formatter = new DataFormatter();
                List<List<String>> data = new ArrayList<>();

                for (Row row : sheet){
                    List<String> cells = new ArrayList<>();
                    int lastCell = row.getLastCellNum() == -1 ? 0 : row.getLastCellNum();
                    
                    for (int i=0; i<lastCell; i++){
                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        cells.add(formatter.formatCellValue(cell));
                    }
                    data.add(cells);
                }

                List<String> headers = new ArrayList<>();
                List<List<String>> rows = new ArrayList<>();
                if(!data.isEmpty()){
                    headers = data.get(0);
                    rows = data.size() > 1 ? data.subList(1, data.size()) : List.of();
                }

                return new ExcelParseRsponse(sheetName, headers, rows);
            }
    }
}
