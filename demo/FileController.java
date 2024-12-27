package com.example.demo;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

@CrossOrigin(origins = "http://localhost:63342")
@RestController
@RequestMapping("/files")
public class FileController {

    @PostMapping("/uploadZip")
    public Map<String, Map<String, Object>> uploadAndCalculateFromZip(@RequestParam("file") MultipartFile zipFile) {
        Map<String, Map<String, Object>> accountResults = new HashMap<>();

        try (InputStream zipInputStream = zipFile.getInputStream();
             ZipInputStream zis = new ZipInputStream(zipInputStream)) {

            ZipEntry entry;
            while ((entry = zis.getNextEntry()) != null) {
                if (!entry.getName().endsWith(".xlsx")) {
                    continue; // Skip non-Excel files
                }

                // Use a ByteArrayOutputStream to hold the content of the current file
                ByteArrayOutputStream buffer = new ByteArrayOutputStream();
                byte[] tempBuffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = zis.read(tempBuffer)) != -1) {
                    buffer.write(tempBuffer, 0, bytesRead);
                }
                InputStream excelStream = new ByteArrayInputStream(buffer.toByteArray());

                Workbook workbook = WorkbookFactory.create(excelStream);
                int numberOfSheets = workbook.getNumberOfSheets();

                Map<String, Object> results = new HashMap<>();
                if (numberOfSheets == 1) {
                    Sheet singleSheet = workbook.getSheetAt(0);
                    PostPaidSheetProcessor.process(singleSheet, results);
                } else if (numberOfSheets >= 2) {
                    Sheet mdnSheet = workbook.getSheetAt(0);
                    PostPaidSheetProcessor.process(mdnSheet, results);

                    Sheet mrSheet = workbook.getSheetAt(1);
                    SeparateMRSheetProcessor.process(mrSheet, results);
                }

                // Store results by account (Excel file name)
                String accountName = entry.getName().replace(".xlsx", ""); // Remove extension for clarity
                accountResults.put(accountName, results);

                workbook.close();
            }
        } catch (Exception e) {
            Map<String, Object> error = new HashMap<>();
            error.put("error", "Error processing ZIP file: " + e.getMessage());
            accountResults.put("Error", error);
            return accountResults;
        }


        return accountResults;
    }



}
