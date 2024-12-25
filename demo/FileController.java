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

            int numberOfSheets = workbook.getNumberOfSheets();

            if (numberOfSheets == 1) {
                // Single sheet with MDN and MR in one column
                Sheet singleSheet = workbook.getSheetAt(0);
                Map<String, Object> singleSheetResults = new HashMap<>();
                PostPaidSheetProcessor.process(singleSheet, singleSheetResults);

                totalIncomeMDN += (double) singleSheetResults.getOrDefault("Total Income MDN", 0.0);
                countMDN += (int) singleSheetResults.getOrDefault("Num of MDNs", 0);
                totalIncomeMR += (double) singleSheetResults.getOrDefault("Total Income MR", 0.0);
                countMR += (int) singleSheetResults.getOrDefault("Num of MRs", 0);
            } else if (numberOfSheets >= 2) {
                // Process MDN sheet
                Sheet mdnSheet = workbook.getSheetAt(0);
                Map<String, Object> mdnResults = new HashMap<>();
                PostPaidSheetProcessor.process(mdnSheet, mdnResults);

                totalIncomeMDN += (double) mdnResults.getOrDefault("Total Income MDN", 0.0);
                countMDN += (int) mdnResults.getOrDefault("Num of MDNs", 0);

                // Process MR sheet
                Sheet mrSheet = workbook.getSheetAt(1);
                Map<String, Object> mrResults = new HashMap<>();
                SeparateMRSheetProcessor.process(mrSheet, mrResults);

                totalIncomeMR += (double) mrResults.getOrDefault("Total Income MR", 0.0);
                countMR += (int) mrResults.getOrDefault("Num of MRs", 0);
            }

            workbook.close();
        } catch (Exception e) {
            results.put("error", "Error processing file: " + e.getMessage());
            return results;
        }

        results.put("Total Income MDN", totalIncomeMDN);
        results.put("Num of MDNs", countMDN);
        results.put("Total Income MR", totalIncomeMR);
        results.put("Num of MRs", countMR);
        return results;
    }
}
