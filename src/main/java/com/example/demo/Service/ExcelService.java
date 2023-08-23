// ExcelService.java
package com.example.demo.Service;

import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.Model.ExcelRow;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Service
public class ExcelService {

    public List<ExcelRow> processExcelFile(MultipartFile file) throws IOException {
        List<ExcelRow> resultList = new ArrayList<>();
        Workbook workbook = WorkbookFactory.create(file.getInputStream());

        Sheet sheet = workbook.getSheetAt(0); // Assumindo que estamos trabalhando com a primeira planilha

        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // Pule a primeira linha (cabeçalho)

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Cell cellA = row.getCell(1);
            Cell cellB = row.getCell(2);

            ExcelRow excelRow = new ExcelRow();
            excelRow.setColumnA(cellA.getNumericCellValue());
            excelRow.setColumnB(cellB.getNumericCellValue());

            // Lógica de comparação e geração da coluna de resultado
            // Pode ser implementada aqui, por exemplo:
            Number resultValue = compareAndGenerateResult(excelRow.getColumnA(), excelRow.getColumnB());
            excelRow.setResultColumn(resultValue);

            resultList.add(excelRow);
        }

        workbook.close();
        return resultList;
    }

    private Number compareAndGenerateResult(Number valueA, Number valueB) {
        // Lógica de comparação e geração do resultado
        // Implemente de acordo com suas necessidades
        // Por exemplo:
        if (valueA.equals(valueB)) {
            return valueA;
        } else {
            return valueB;
        }
    }
}
