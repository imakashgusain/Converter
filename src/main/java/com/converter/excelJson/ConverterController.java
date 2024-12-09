package com.converter.excelJson;

import com.converter.excelJson.util.ExcelUtils;

import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.List;

@RestController
@RequestMapping("/api/convert")
public class ConverterController {

    @PostMapping("/excel-to-json")
    public ResponseEntity<List<List<String>>> convertExcelToJson(@RequestParam("file") MultipartFile file) {
        try {
            // Convert MultipartFile to File
            File tempFile = File.createTempFile("uploaded-", file.getOriginalFilename());
            file.transferTo(tempFile);

            // Use ExcelUtils to read the Excel file and convert it to a List of rows (JSON structure)
            List<List<String>> excelData = ExcelUtils.readExcelFile(tempFile);

            return new ResponseEntity<>(excelData, HttpStatus.OK);
        } catch (IOException e) {
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @PostMapping("/json-to-excel")
    public ResponseEntity<byte[]> convertJsonToExcel(@RequestBody List<List<String>> jsonData) {
        try {
            // Create Excel file from the provided JSON data
            File excelFile = ExcelUtils.createExcelFileFromJson(jsonData, "converted");

            // Return the Excel file as a response
            byte[] excelBytes = java.nio.file.Files.readAllBytes(excelFile.toPath());
            HttpHeaders headers = new HttpHeaders();
            headers.add("Content-Disposition", "attachment; filename=converted.xlsx");

            return ResponseEntity.ok().headers(headers).body(excelBytes);
        } catch (IOException e) {
            return ResponseEntity.status(500).build();
        }
    }
}
