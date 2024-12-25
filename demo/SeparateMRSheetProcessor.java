// SeparateMRSheetProcessor.java
package com.example.demo;

import org.apache.poi.ss.usermodel.*;

import java.util.Map;
import java.util.regex.Pattern;

public class SeparateMRSheetProcessor {

    public static void process(Sheet sheet, Map<String, Object> results) {
        double totalIncomeMR = 0.0;
        int countMR = 0;

        int iccidColumnIndex = -1;
        int totalColumnIndex = -1;

        // Define the MR pattern
        Pattern mrPattern = Pattern.compile("MR40-A1-0000\\d{4}");

        Row headerRow = sheet.getRow(3); // Assuming header is at row 3
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String cellValue = cell.toString().toLowerCase().trim();
                if (cellValue.contains("iccid")) {
                    iccidColumnIndex = cell.getColumnIndex();
                } else if (cell.getCellType() == CellType.STRING || cell.getCellType() == CellType.NUMERIC) {
                    totalColumnIndex = cell.getColumnIndex(); // Dynamically locate the date column
                }
            }
        }

        if (iccidColumnIndex == -1 || totalColumnIndex == -1) {
            results.put("error", "Required columns not found in MR sheet.");
            return;
        }

        FormulaEvaluator evaluator = SheetProcessor.getFormulaEvaluator(sheet);

        for (int i = 4; i <= sheet.getLastRowNum(); i++) { // Start from row 4 (data rows)
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell iccidCell = row.getCell(iccidColumnIndex);
            Cell totalCell = row.getCell(totalColumnIndex);

            if (iccidCell == null || totalCell == null) continue;

            String iccid = iccidCell.toString().trim();
            double total = evaluator.evaluate(totalCell).getNumberValue();

            if (mrPattern.matcher(iccid).matches()) {
                totalIncomeMR += total;
                countMR++;
            }
        }

        results.put("Num of MRs", countMR);
        results.put("Total Income MR", totalIncomeMR);
    }
}
