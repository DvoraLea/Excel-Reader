package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

@RestController
@RequestMapping("/files")
public class FileController {

    @PostMapping("/upload")
    public Map<String, Object> uploadAndCalculate(@RequestParam("file") MultipartFile file) {
        Map<String, Object> results = new HashMap<>();
        double totalIncomeMDN = 0.0;
        double totalIncomeMR = 0.0;
        int countMDN = 0;
        int countMR = 0;

        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Process the first sheet (PostPaid/Sheet1)
            Sheet sheet1 = workbook.getSheetAt(0); // Assuming first sheet is PostPaid or Sheet1
            if (sheet1 != null) {
                Map<String, Object> postPaidResults = new HashMap<>();
                processPostPaidSheet(sheet1, postPaidResults);

                // Update totals and counts for MDN
                totalIncomeMDN += (double) postPaidResults.getOrDefault("Total Income MDN", 0.0);
                countMDN += (int) postPaidResults.getOrDefault("Num of MDNs", 0);

                // Update totals and counts for MR
                totalIncomeMR += (double) postPaidResults.getOrDefault("Total Income MR", 0.0);
                countMR += (int) postPaidResults.getOrDefault("Num of MRs", 0);
            }

            // Process the second sheet (Netrifi/MR Sheet for Mad Bee)
            if (workbook.getNumberOfSheets() > 1) {
                Sheet sheet2 = workbook.getSheetAt(1); // Assuming second sheet is Netrifi or MR sheet
                if (sheet2 != null) {
                    Map<String, Object> netrifiResults = new HashMap<>();
                    processNetrifiSheet(sheet2, netrifiResults);

                    // Update totals and counts for MR
                    totalIncomeMR += (double) netrifiResults.getOrDefault("Total Income MR", 0.0);
                    countMR += (int) netrifiResults.getOrDefault("Num of MRs", 0);
                }
            }

            workbook.close();

        } catch (Exception e) {
            results.put("error", "Error processing file: " + e.getMessage());
            return results;
        }

        // Combine and return the results
        results.put("Total Income MDN", totalIncomeMDN);
        results.put("Num of MDNs", countMDN);
        results.put("Total Income MR", totalIncomeMR);
        results.put("Num of MRs", countMR);

        return results;
    }

    private void processPostPaidSheet(Sheet sheet, Map<String, Object> results) {
        double totalIncomeMDN = 0.0;
        double totalIncomeMR = 0.0;
        int countMDN = 0;
        int countMR = 0;

        int phoneNumberColumnIndex = -1;
        int totalAfterFeeColumnIndex = -1;

        // Locate header row to identify column indices
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

        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        // Process each row
        for (Row row : sheet) {
            if (row.getRowNum() == 0) continue; // Skip header row

            Cell phoneNumberCell = row.getCell(phoneNumberColumnIndex);
            Cell totalAfterFeeCell = row.getCell(totalAfterFeeColumnIndex);

            if (phoneNumberCell == null || totalAfterFeeCell == null) continue;

            String phoneNumber = "";
            if (phoneNumberCell.getCellType() == CellType.NUMERIC) {
                // Convert numeric phone number to string
                phoneNumber = String.valueOf((long) phoneNumberCell.getNumericCellValue());
            } else {
                phoneNumber = phoneNumberCell.toString().trim();
            }

            double total = 0.0;

            // Handle formulas and numeric cells
            if (totalAfterFeeCell.getCellType() == CellType.NUMERIC) {
                total = totalAfterFeeCell.getNumericCellValue();
            } else if (totalAfterFeeCell.getCellType() == CellType.FORMULA) {
                total = evaluator.evaluate(totalAfterFeeCell).getNumberValue();
            }

            // Determine if it's an MDN or MR
            if (phoneNumber.matches("\\d+")) { // MDN: Only digits
                totalIncomeMDN += total;
                countMDN++;
            } else if (phoneNumber.startsWith("MR")) { // MR: Starts with "MR"
                totalIncomeMR += total;
                countMR++;
            }
        }

        results.put("Num of MDNs", countMDN);
        results.put("Total Income MDN", totalIncomeMDN);
        results.put("Num of MRs", countMR);
        results.put("Total Income MR", totalIncomeMR);
    }

    private void processNetrifiSheet(Sheet sheet, Map<String, Object> results) {
        double totalIncomeMR = 0.0;
        int countMR = 0;

        int iccidColumnIndex = -1;
        int totalColumnIndex = -1;

        // Locate the ICCID and total columns
        Row headerRow = sheet.getRow(3); // Start from row 4 (index 3)
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String cellValue = cell.toString().toLowerCase().trim();
                if (cellValue.contains("iccid")) {
                    iccidColumnIndex = cell.getColumnIndex();
                } else if (cellValue.matches("\\d{1,2}/\\d{1,2}/\\d{4}")) { // Matches a date-like header
                    totalColumnIndex = cell.getColumnIndex();
                }
            }
        }

        if (iccidColumnIndex == -1 || totalColumnIndex == -1) {
            results.put("error", "Required columns not found in Netrifi sheet.");
            return;
        }

        FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();

        // Start processing from row 4 (index 3)
        for (int i = 4; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell iccidCell = row.getCell(iccidColumnIndex);
            Cell totalCell = row.getCell(totalColumnIndex);

            if (iccidCell == null || totalCell == null) continue;

            String iccid = iccidCell.toString().trim();
            double total = 0.0;

            // Handle formulas and numeric cells
            if (totalCell.getCellType() == CellType.NUMERIC) {
                total = totalCell.getNumericCellValue();
            } else if (totalCell.getCellType() == CellType.FORMULA) {
                total = evaluator.evaluate(totalCell).getNumberValue();
            }

            if (iccid.startsWith("MR")) {
                totalIncomeMR += total;
                countMR++;
            }
        }

        results.put("Num of MRs", countMR);
        results.put("Total Income MR", totalIncomeMR);
    }
}
