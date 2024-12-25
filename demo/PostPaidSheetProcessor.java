// Updated PostPaidSheetProcessor.java
package com.example.demo;

import org.apache.poi.ss.usermodel.*;

import java.util.Map;
import java.util.regex.Pattern;

public class PostPaidSheetProcessor {

    public static void process(Sheet sheet, Map<String, Object> results) {
        double totalIncomeMDN = 0.0;
        double totalIncomeMR = 0.0;
        int countMDN = 0;
        int countMR = 0;

        int phoneNumberColumnIndex = -1;
        int totalAfterFeeColumnIndex = -1;

        // Define patterns
        Pattern mrPattern = Pattern.compile("MR40-A1-0000\\d{4}");
        Pattern mdnPattern = Pattern.compile("\\d{10}"); // Assuming MDNs are 10-digit numbers

        Row headerRow = sheet.getRow(0);
        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue().toLowerCase().trim();
            if (header.contains("phone number")) {
                phoneNumberColumnIndex = cell.getColumnIndex();
            } else if (header.contains("total")) {
                totalAfterFeeColumnIndex = cell.getColumnIndex();
            }
        }

        if (phoneNumberColumnIndex == -1 || totalAfterFeeColumnIndex == -1) {
            results.put("error", "Required columns not found in PostPaid sheet.");
            return;
        }

        FormulaEvaluator evaluator = SheetProcessor.getFormulaEvaluator(sheet);

        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue;

            Cell phoneNumberCell = row.getCell(phoneNumberColumnIndex);
            Cell totalAfterFeeCell = row.getCell(totalAfterFeeColumnIndex);

            if (phoneNumberCell == null || totalAfterFeeCell == null) continue;

            String phoneNumber = phoneNumberCell.toString().trim();
            double total = evaluator.evaluate(totalAfterFeeCell).getNumberValue();

            // Handle MR pattern
            if (mrPattern.matcher(phoneNumber).matches()) {
                totalIncomeMR += total;
                countMR++;
            }
            // Handle MDN pattern (numeric phone numbers)
            else if (mdnPattern.matcher(phoneNumber).matches()) {
                totalIncomeMDN += total;
                countMDN++;
            }
            // Handle numbers in scientific notation
            else if (phoneNumber.matches("\\d\\.\\d+E\\d+")) {
                String convertedNumber = String.format("%.0f", Double.parseDouble(phoneNumber));
                if (mdnPattern.matcher(convertedNumber).matches()) {
                    totalIncomeMDN += total;
                    countMDN++;
                } else {
                    System.out.println("Uncategorized row (scientific notation): " + phoneNumber + " -> " + convertedNumber);
                }
            }
            // Uncategorized
            else {
                System.out.println("Uncategorized row: " + phoneNumber);
            }
        }

        results.put("Num of MDNs", countMDN);
        results.put("Total Income MDN", totalIncomeMDN);
        results.put("Num of MRs", countMR);
        results.put("Total Income MR", totalIncomeMR);
    }
}
